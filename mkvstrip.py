#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Copyright: (c) 2013 William Forde (willforde@gmail.com)
# License: GPLv3, see LICENSE for more details
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <http://www.gnu.org/licenses/>.


__version__ = "1.0.0"

from functools import lru_cache
from operator import itemgetter
from pymediainfo import MediaInfo
import subprocess
import argparse
import itertools
import time
import json
import sys
import os

# Global parser namespace
cli_args = None

MKVMERGE_DEFAULT = "mkvmerge"
MEDIAINFO_DEFAULT = "mediainfo"
MKVPROPEDIT_DEFAULT = "mkvpropedit"


def catch_interrupt(func):
    """Decorator to catch Keyboard Interrupts and silently exit."""
    def wrapper(*args, **kwargs):
        try:
            func(*args, **kwargs)
        except KeyboardInterrupt:  # pragma: no cover
            pass

    # The function been catched
    return wrapper


def walk_directory(path):
    """
    Walk through the given directory to find all mkv files and process them.

    :param str path: Path to Directory containing mkv files.

    :return: List of processed mkv files.
    :rtype: list[str]
    """
    movie_list = []
    if os.path.isfile(path):
        if path.lower().endswith(".mkv"):
            movie_list.append(path)
        else:
            raise ValueError("Given file is not a valid mkv file: '%s'" % path)

    elif os.path.isdir(path):
        dirs = []
        # Walk through the directory
        for dirpath, _, filenames in os.walk(path):
            files = []
            for filename in filenames:
                if filename.lower().endswith(".mkv"):
                    files.append(filename)

            # Sort list of files and add to directory list
            dirs.append((dirpath, sorted(files)))

        # Sort the list of directorys & files and process them
        for dirpath, filenames in sorted(dirs, key=itemgetter(0)):
            for filename in filenames:
                fullpath = os.path.join(dirpath, filename)
                movie_list.append(fullpath)
    else:
        raise FileNotFoundError("[Errno 2] No such file or directory: '%s'" % path)

    return movie_list


def edit_file(command):
    """
    Edit a mkv file with the given parameters.

    :param list command: The list of command parameters to pass to the editing app.

    :return: Boolean indicating if edit was successful.
    :rtype: bool
    """
    # Skip editing if in dry run mode
    if cli_args.verbose:
        print(command)
    
    if cli_args.dry_run:
        print("Dry run 100%")
        return False
    
    sys.stdout.write("Progress 0%")
    sys.stdout.flush()

    try:
        # Call subprocess command to edit file
        process = subprocess.Popen(command, stdout=subprocess.PIPE, universal_newlines=True)

        # Display Percentage until subprocess has finished
        retcode = process.poll()
        while retcode is None:
            # Sleep for a quarter second and then dislay progress
            time.sleep(.25)
            for line in iter(process.stdout.readline, ""):
                if "progress" in line.lower():
                    sys.stdout.write("\r%s" % line.strip())
                    sys.stdout.flush()

            # Check return code of subprocess
            retcode = process.poll()

        # Check if return code indicates an error
        sys.stdout.write("\n")
        if retcode:
            raise subprocess.CalledProcessError(retcode, command, output=process.stdout)

    except subprocess.CalledProcessError as e:
        print("Remux failed!")
        print(e)
        return False
    else:
        return True


def replace_file(tmp_file, org_file):
    """
    Replaces the original mkv file with the newly remuxed temp file.

    :param str tmp_file: The temporary mkv file
    :param str org_file: The original mkv file to replace.
    """
    # Preserve timestamp
    stat = os.stat(org_file)
    os.utime(tmp_file, (stat.st_atime, stat.st_mtime))

    # Overwrite original file
    try:
        os.unlink(org_file)
        os.rename(tmp_file, org_file)
    except EnvironmentError as e:
        os.unlink(tmp_file)
        print("Renaming failed: %s => %s" % (tmp_file, org_file))
        print(e)


class AppendSplitter(argparse.Action):
    """
    Custom action to split multiple parameters which are
    separated by a comma, and append then to a default list.
    """
    def __call__(self, _, namespace, values, option_string=None):
        items = self.default if isinstance(self.default, list) else []
        items.extend(values.split(","))
        setattr(namespace, self.dest, items)


class RealPath(argparse.Action):
    """
    Custom action to convert given path to a full canonical path,
    eliminating any symbolic links if encountered.
    """
    def __call__(self, _, namespace, value, option_string=None):
        setattr(namespace, self.dest, os.path.realpath(value))


class MKVFile(object):
    """
    Extracts track information contained within a Matroska file and
    checks for unwanted audio & subtitle tracks.

    :param str path: Path to the Matroska file to process.
    """
    def __init__(self, path):
        self.dirpath, self.filename = os.path.split(path)
        self.path = path
               
        media_info = MediaInfo.parse(path)
        self.general_tracks = media_info.general_tracks
        self.video_tracks = media_info.video_tracks
        self.audio_tracks = media_info.audio_tracks
        self.subtitle_tracks = media_info.text_tracks
        self.menu_tracks = media_info.menu_tracks
        self.streamorder_video = []
        self.streamorder_audio = []
        self.streamorder_subtitles = []
        self.track_order = []
        self.subtitles_forced = []
        self.streams_misaligned = False

    @lru_cache()
    def _filtered_tracks(self, track_type):
        """
        Return a tuple consisting of tracks to keep and tracks to remove, if
        there are indeed tracks that need to be removed, else return False.

        Available track types:
            subtitle
            audio

        :param str track_type: The track type to check.

        :return: Tuple of tracks to keep and remove
        :rtype: tuple[list[Track]]
        """
        if track_type == 'Audio':
            languages_to_keep = cli_args.language
            tracks = self.audio_tracks
        elif track_type == 'Text':
            languages_to_keep = cli_args.sub_language
            tracks = self.subtitle_tracks
            
        # Lists of track to keep & remove
        remove = []
        keep = []
        # Iterate over all tracks to find which track to keep or remove
        for track in tracks:
            if track.language in languages_to_keep:
                # Tracks we want to keep               
                if cli_args.sub_forced and track_type == "Text":
                    if track.forced == "Yes":
                        keep.append(track)
                    else:
                        remove.append(track)              
                else:
                    keep.append(track)
            else:
                # Tracks we want to remove
                remove.append(track)
        return keep, remove

    @property
    def remux_required(self):
        """
        Check if any remuxing of the mkv files is required.

        :return: Return True if remuxing is required else False
        :rtype: bool
        """

        audio_to_keep, audio_to_remove = self._filtered_tracks("Audio")
        sub_to_keep, sub_to_remove = self._filtered_tracks("Text")
              
        for track in self.video_tracks:
            self.streamorder_video.append(track.streamorder)

        for track in self.audio_tracks:
            self.streamorder_audio.append(track.streamorder)

        for track in self.subtitle_tracks:
            self.streamorder_subtitles.append(track.streamorder)
        
        for video, audio in itertools.product(self.streamorder_video, self.streamorder_audio):
            if video > audio:
                self.streams_misaligned = True
                print("Misaligned streams detected")
                break
        
        if not self.streams_misaligned:
            for video, subtitles in itertools.product(self.streamorder_video, self.streamorder_subtitles):
                if video > subtitles:
                    self.streams_misaligned = True
                    print("Misaligned streams detected")
                    break
        
        if not self.streams_misaligned:
            for audio, subtitles in itertools.product(self.streamorder_audio, self.streamorder_subtitles):
                if audio > subtitles:
                    self.streams_misaligned = True
                    print("Misaligned streams detected")
                    break

        has_something_to_remove = audio_to_remove or sub_to_remove
        if has_something_to_remove or self.streams_misaligned:
            return True
        else:
            return False

    def remove_tracks(self):
        """Remove the unwanted tracks."""

        command = [cli_args.mkvmerge, "--output"]
        print("\nRemuxing:", self.filename)
        print("============================")

        # Output the remuxed file to a temp tile, This will protect
        # the original file from been corrupted if anything goes wrong
        if cli_args.tmp_dir:
           tmp_path_real = os.path.realpath(cli_args.tmp_dir)
           print(tmp_path_real)
           tmp_file = u"%s/%s.tmp" % (tmp_path_real, self.filename)
           print(tmp_file)
        else:    
            tmp_file = u"%s.tmp" % self.path
        
        command.append(tmp_file)
        
        command.extend(["--title", self.filename[:(self.filename.index("[")-1)]])
        command.extend(["--no-chapters"])
        command.extend(["--no-attachments"])
        command.extend(["--no-track-tags"])
        command.extend(["--disable-track-statistics-tags"])
        
        for track in self.video_tracks:
            command.extend(["--track-name", ":".join((str(track.streamorder)," "))])
            command.extend(["--language", ":".join((str(track.streamorder),"und"))])
            self.track_order.extend(str(track.streamorder))
                
        # Iterate over all tracks and mark which tracks are to be kept
        for track_type in ("Audio", "Text"):
            keep, remove = self._filtered_tracks(track_type)
            sorted_keep =[]
            keep_ids = []
            
            if track_type == "Audio":
                for lang in cli_args.language:
                    internal_keep = []
                    for track in keep:
                        if track.language == lang:
                            internal_keep.append(track)
                    internal_keep.sort(key=lambda x: x.stream_size, reverse=True)
                    sorted_keep.extend(internal_keep)
            
            if track_type == "Text":
                for lang in cli_args.sub_language:
                    internal_keep = []
                    for track in keep:
                        if track.language == lang:
                            internal_keep.append(track)
                    internal_keep.sort(key=lambda x: x.forced, reverse=True)
                    sorted_keep.extend(internal_keep)

            print("Retaining %s track(s):" % track_type)
            for count, track in enumerate(sorted_keep):
                keep_ids.append(str(track.streamorder))
                print("   ", "Track #{}: {} - {}".format(track.streamorder, track.language, track.format))

                # Set the first track as default
                command.extend(["--default-track", ":".join((str(track.streamorder), "0" if count else "1"))])
            
                #Set Track names
                if track_type == "Audio":
                    command.extend(["--track-name", ":".join((str(track.streamorder),track.commercial_name))])
                
                elif track_type == "Text":
                    command.extend(["--track-name", ":".join((str(track.streamorder),"{}{}".format(track.other_language[0], " [Forced]" if track.forced == "Yes" else "")))])
                
            # Set which tracks are to be kept
            if keep_ids and track_type == "Audio":
                command.extend(["--audio-tracks", ",".join(keep_ids)])
                self.track_order.extend(keep_ids)
            elif keep_ids and track_type == "Text":
                command.extend(["--subtitle-tracks", ",".join(keep_ids)])
                self.track_order.extend(keep_ids)
            elif track_type == "Text" and not keep_ids:
                command.extend(["--no-subtitles"])
            elif track_type == "Audio" and not keep_ids:
                command.extend(["--no-audio"])
                 
            print("Removing %s track(s):" % track_type)
            for track in remove:
                print("   ", "Track #{}: {} - {}".format(track.streamorder, track.language, track.format))

            print("----------------------------")

        command.extend(["--track-order", "{1}{0}".format(",0:".join(self.track_order), "0:")])

        # Add source mkv file to command and remux
        command.append(self.path)
        if edit_file(command):
            replace_file(tmp_file, self.path)
        else:
            if os.path.exists(tmp_file):
                os.remove(tmp_file)
                raise Exception("Remuxing failed, but the file on disk should be OK.")   
    
    def cleanup(self):
        command = [cli_args.mkvpropedit, self.path]
        print("\nCleaning Up:", self.filename)
        print("============================")
        

        for track in self.video_tracks:
            if track.title:
                command.extend(["--edit", ":".join(("track",str(track.track_id))), "--delete", "name"])
            if track.language:
                command.extend(["--edit", ":".join(("track",str(track.track_id))), "--set", "language=und"])
        
        for track in self.audio_tracks:
            if track.title != track.commercial_name:
                command.extend(["--edit", ":".join(("track",str(track.track_id))), "--set", "=".join(("name",track.commercial_name))])
        
        for track in self.subtitle_tracks:
            if track.title != "{}{}".format(track.other_language[0], " [Forced]" if track.forced == "Yes" else ""):
                command.extend(["--edit", ":".join(("track",str(track.track_id))), "--set", "=".join(("name","{}{}".format(track.other_language[0], " [Forced]" if track.forced == "Yes" else "")))])
        
        for track in self.general_tracks:
            if track.attachments:
                attachment_list = track.attachments.split(" / ")
                for attachment in attachment_list:
                    command.extend(["--delete-attachment", ":".join(("name", attachment))])
        
        if self.menu_tracks:
            command.extend(["-c", ""])
            
        
        track_statistics = False 
        while track_statistics == False:
            for track in self.video_tracks:
                if track.duration_source != "General_Duration" and track.framecount_source != "General_Duration":
                    track_statistics = True
            for track in self.audio_tracks:
                if track.duration_source != "General_Duration" and track.samplingcount_source != "General_Duration":
                    track_statistics = True
            break

        if track_statistics == True:
            command.extend(["--delete-track-statistics-tags"])


        if len(command) >= 3:
            if edit_file(command):
               print("Cleaned up {} successfully".format(self.path))
        else:
            print("Nothing to do here")
        

@catch_interrupt
def main(params=None):
    """
    Check all mkv files an remove unnecessary tracks.

    :param params: [opt] List of arguments to pass to argparse.
    :type params: list or tuple
    """
    # Create Parser to parse the required arguments
    parser = argparse.ArgumentParser(description="Strips unwanted tracks from MKV files and cleans them up.")
    parser.add_argument("paths", nargs='+', help="Path to media file(s).")
    parser.add_argument("--mediainfo", action="store", default=MEDIAINFO_DEFAULT, metavar="path", help="Path to the mediainfo binary.")
    parser.add_argument("--mkvmerge", action="store", default=MKVMERGE_DEFAULT, metavar="path", help="Path to the mkvmerge binary.")
    parser.add_argument("--mkvpropedit", action="store", default=MKVPROPEDIT_DEFAULT, metavar="path", help="Path to the mkvpropedit binary.")
    parser.add_argument("--tmp-dir", action="store", default=None, metavar="path", help="Custom Path for temporary files, if it does not exist it is created")
    parser.add_argument("-l", "--language",  action=AppendSplitter, default=None, required=True, metavar="language", help="Comma-separated list of ISO 639-1 compliant language codes defining the audio languages to retain.")
    parser.add_argument("-s", "--sub-language", action=AppendSplitter, default=None, required=True, metavar="language", help="Comma-separated list of ISO 639-1 compliant language codes defining the subtitle languages to retain.")
    parser.add_argument("-f", "--sub-forced", action="store_true", default=False, help="When enabled only forced subtitles are kept.")
    parser.add_argument("-d", "--dry-run", action="store_true", default=False, help="Dry run for testing.")
    parser.add_argument("-v", "--verbose", action="store_true", default=False, help="Verbose output.")

    # Parse the list of given arguments
    globals()["cli_args"] = parser.parse_args(params)

    # Iterate over all found mkv files
    print("Searching for MKV files to process.")
    print("Warning: This may take some time...")
    for path in cli_args.paths:
        path = os.path.realpath(path)
        for mkv_file in walk_directory(path):
            if cli_args.verbose:
                print("Checking", mkv_file)
            mkv_obj = MKVFile(mkv_file)
            if mkv_obj.remux_required:
                mkv_obj.remove_tracks()
            else:
                mkv_obj.cleanup()
            

if __name__ == "__main__":
    main()

