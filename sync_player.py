# -*- coding: utf-8 -*-
"""
    SyncPlayer.py - A Python script for synchronizing music files between 
    Foobar2000 and an MP3 device.

    Copyright (C) 2010 Blair Sutton

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see .  
"""

#
# TODO: embed album art in transcoded files
#requires lame, sed, flac, flacmeta on the path
#

########################
#### SETTINGS BELOW ####
########################

devices = [u"car", u"g2", u"droid3", u"cowon"]

# this is an array of foobar2000 playlist names you want synched
#
playlists_array = [[u"car"], [u"mp3player"], [u"car"], [u"mp3player"]]  

# this is the path to your android/mp3 player music folder once mounted.
# the converted files will be placed here.
#
# Note: expected to start with 'X:\' where X is a drive letter
# [:-1]
destination_root_array = [ u"", u"Music", u"Music", u"Music"]


# this is the path to your android/mp3 player playlist folder. m3u files will
# be placed here.
#
playlist_root_array = [u"Playlists", u"Music\Playlists", u"Music\Playlists", u"Music\Playlists"]



# this is your target conversion format. only mp3 works now
#
destination_ext = u".mp3"

# this is how many path levels of your source dir to ignore
# ie - if you put music in X:\Music\stuff in foobar,
# this should be 2 to make a folder called stuff under
# destination_root defined above
# 
path_ignore_depth = 2
do_convert_files = True


####################
#### CODE BELOW ####
####################

tag_list = ["TITLE","ARTIST","ALBUM","DATE","TRACKNUMBER","GENRE"]

import win32com.client
import codecs
from win32com.client.gencache import EnsureDispatch
from os.path import basename, splitext, exists, join
from os import sep, makedirs, environ, walk, remove
import sys
import unicodedata
from subprocess import call
import shlex, subprocess
from urlparse import urlparse, urlsplit
from urllib import quote, unquote
from logging import StreamHandler, Formatter, DEBUG, INFO, getLogger
from shutil import copyfile
import argparse




log = getLogger("Foobar.MP3PlayerSync")
log.setLevel(DEBUG)
lh = StreamHandler()
lh.setFormatter(Formatter("%(levelname)s|%(asctime)s|%(filename)s:%(lineno)d|%(message)s"))
log.addHandler(lh)

log.info("Connecting to foobar2000 COM server...")

prog_id = "Foobar2000.Application.0.7"

fb2k = win32com.client.Dispatch(prog_id)
all_files = []
failed_files = []
file_errors = 0
files_copied = 0
files_transcoded = 0
files_skipped = 0
files_deleted = 0
file_remove_errors = 0
playlists = u""
playlist_root = u""
destination_root = u""


def main():
    select_device()
    select_drive_letter()
    log.info("Scanning drive for existing files...")
    global files_deleted, file_remove_errors, all_files 
    all_files = scan_dir(destination_root, [destination_ext])
    log.info("Found %i files on the device." % len(all_files))
    if confirm("Would you like to print them?", False):
      for each_file in all_files:
        print each_file.encode(sys.getfilesystemencoding())
      confirm("Press enter.", True)
    
    fb2k_playlists = [ i for i in fb2k.Playlists if i.Name in playlists ]
    if fb2k_playlists:
        for pl in fb2k_playlists:
            sync_playlist(pl)
    log.info("After sync %i files not matched to a playlist." % len(all_files))
    if confirm("Print them?", True):
      for each_file in all_files:      
        print each_file.encode(sys.getfilesystemencoding())
    if len(all_files) > 0:
      if confirm("Delete these " + str(len(all_files)) + " files?", False):      
        for each_file in all_files:
          try:
            log.info("Removing: %s" % each_file.encode(sys.getfilesystemencoding()))
            remove(each_file.encode(sys.getfilesystemencoding()))
          except WindowsError:
            log.error("Error removing file: %s" % each_file.encode(sys.getfilesystemencoding()))
            file_remove_errors += 1
            continue
          
        files_deleted += 1
        
    log.info("Completed Sync!")
    log.info("Copied %i files" % files_copied )
    log.info("Transcoded %i files" % files_transcoded )
    log.info("Skipped %i files" % files_skipped )
    log.info("Deleted %i files" % files_deleted )
    log.info("%i files not copied due to errors" % len(failed_files) )
    log.info("%i files not deleted due to error" % file_remove_errors )
    if len(failed_files) > 0:
      if confirm("Print files that didn't copy?", True):
        for each_file in failed_files:    
          log.info("%s" % each_file.encode(sys.getfilesystemencoding()) )
    
    
    
def scan_dir(path, exts):
    """
    path    -    where to begin folder scan
    """
    selected_files = []
    
    for root, dirs, files in walk(path): 
        selected_files += select_files(root, files, exts)
            
    return selected_files

def whatisthis(s):
    if isinstance(s, str):
      print "ordinary string"
    elif isinstance(s, unicode):
      print "unicode string"
    else:
      print "not a string"


    
def select_files(root, files, exts):
    """
    simple logic here to filter out interesting files
    """
    selected_files = []
    for my_file in files:
        #do concatenation here to get full path            
        full_path = join(root, my_file)
        file_ext = splitext(my_file)[1]        
        for ext in exts:
          if file_ext.lower() == ext.lower():
            selected_files.append(full_path.lower())
          #else:
           # print "Extensions didn't match (" + ext + " f:" + file_ext + "), skipping: " + full_path

    return selected_files   
    

def sync_playlist(sync_playlist):
    log.info("Syncing playlist '%s'..." % sync_playlist.Name)
    tracks = sync_playlist.GetTracks()
    global file_errors, all_files
    m3u_lines = ["#EXTM3U"]
    for t in tracks: 
        try:        
        
          if t.Path[:7] != "file://":
            log.error("Unsupported scheme in playlist, expected file, got: %s" % t.Path[:7])
            file_errors += 1
            failed_files.append(t.Path)
            continue
          source_path = t.Path[7:]              
        
          dest_path = sync_file(source_path)
        except Exception, e:        
          log.error("Error copying track with source path %s: %s" % (t.Path, str(e)))
          file_errors += 1
          failed_files.append(t.Path)
          continue
          
        m3u_lines.append(t.FormatTitle("#EXTINF:%length_seconds%, %artist% - %title%"))
        m3u_lines.append(dest_path[3:])
        #print "adding line to m3u: " + dest_path.encode(sys.getfilesystemencoding()) 
        idx = 0
        try:
          idx = all_files.index(dest_path.lower())
        except (IndexError, ValueError):
          #print "Not in the list: " + dest_path.encode(sys.getfilesystemencoding()) 
          continue
        #print "Found: " + dest_path.encode(sys.getfilesystemencoding()) + " at index " + str(idx)        
        del all_files[idx]
    
    create_m3u(sync_playlist.Name, m3u_lines)
        
def sync_file(source_path):
    global files_transcoded    
    global files_copied
    global files_skipped
    parts_all = source_path.split(sep)
    length = len(parts_all)
    parts = parts_all[-length:]

    filenameext = parts[length-1]
    (filename, ext) = splitext(filenameext)
    
    actual_destination_root = destination_root
    if destination_root[-1] == "\ "[:-1]:
      actual_destination_root = destination_root[:-1]
    
    parts_new_path = [actual_destination_root] 
    parts_new_path.extend(parts[path_ignore_depth:length-1])
    destination_folder = sep.join(parts_new_path)
    parts_new_path.append(filename + destination_ext)     
    destination_path = sep.join(parts_new_path)
    
    
    if not exists(destination_folder):
        log.debug("Creating folder: '%s'..." % destination_folder)
        makedirs(destination_folder)
    
    if not exists(destination_path):
      if (destination_ext.lower() != ext.lower()) and do_convert_files:
        if ext.lower() == ".flac":
          convert_flac_file(source_path, destination_path)
          files_transcoded += 1
        else:
          log.warn("Unsupported extension on file: %s" % source_path)
          files_skipped += 1
      else:
        log.info("Copying: '%s' -> '%s'" % (source_path, destination_path))
        copyfile(source_path, destination_path)
        files_copied += 1
    else:
      files_skipped += 1
        
    return destination_path
    
def get_flac_metadata(input_file, tag_name):
    
    command = "metaflac --show-tag=" + tag_name + " \"" + input_file + "\" | sed s/.*=//"
	
    #use file system encoding when using popen
    proc=subprocess.Popen(command.encode(sys.getfilesystemencoding()), shell=True, stdout=subprocess.PIPE, )
    #use console encoding when reading from console
    output = proc.communicate()[0].decode(sys.stdout.encoding)
	
    #delete annoying line ends
    ret = output.replace('\n','').replace('\r', '')
    
    #replace unicode string chars with similar
    ret = unicodedata.normalize('NFKD', unicode(ret)).encode('ascii','replace')
 
    #print "GOTS ME A " + tag_name + ", ITS " + ret
    return ret
    
    
def convert_flac_file(input_file, output_file):

    
    log.info("Transcoding: '%s' -> '%s'" % (input_file, output_file))
    
    TITLE= "\"" + get_flac_metadata(input_file, "TITLE") + "\""
    ARTIST="\"" + get_flac_metadata(input_file, "ARTIST") + "\""
    ALBUM="\"" + get_flac_metadata(input_file, "ALBUM") + "\""
    DATE="\"" + get_flac_metadata(input_file, "DATE") + "\""
    TRACKNUMBER="\"" + get_flac_metadata(input_file, "TRACKNUMBER") + "\""
    GENRE="\"" + get_flac_metadata(input_file, "GENRE") + "\""
    

    #command = """"%s" -i "%s" -ac 2 -aq 0 -map_metadata 0 "%s" """ % (ffmpeg_path, input_file, output_file)
    command = "flac \"" + input_file + "\" -cd  | lame -V0 --add-id3v2 --tt " + TITLE +\
      " --ta " + ARTIST + " --tl " + ALBUM + " --ty " + DATE + " --tn " +\
      TRACKNUMBER + " --tg " + GENRE + " - \"" + output_file + "\""


    global failed_files, file_errors

    #TODO: call encode once? im lazy, eh
    log.debug("Converter command line:\n%s" % command.encode(sys.getfilesystemencoding()))
    try:
        retcode = call(command.encode(sys.getfilesystemencoding()), shell=True)        
    except OSError, e:
        log.critical("Converter execution failed: %s", e.strerror)
        failed_files.append(input_file)
        file_errors += 1
 
def create_m3u(playlist_name, m3u_lines):
    if not exists(playlist_root):
        log.info("Creating folder: '%s'..." % playlist_root)
        makedirs(playlist_root)
        
    m3u_path = "%s\\%s.m3u" % (playlist_root, playlist_name)
    log.info("Creating m3u playlist: '%s'..." % m3u_path)    
    f = codecs.open(m3u_path, 'w', 'mbcs')
    f.write("\n".join(m3u_lines))
    f.close()
    
def select_device():
    global playlists, destination_root, playlist_root
    print "Which device would you like to sync: "
    for device in devices:
      print device
    prompt = "Select a device name:"
    prompt = '%s [%s]|: ' % (prompt, devices[0])
        
    index = 0    
    while True:    
      ans = raw_input(prompt)
      if ans == "":
        index = 0
        break
      else:
        try:
          index = devices.index(ans.lower())
        except (IndexError, ValueError):
          index = -1      
        if index == -1:
          print 'Please enter a valid device name.'
          continue
        else:
          break
    
    playlists = playlists_array[index]
    destination_root = destination_root_array[index]
    playlist_root = playlist_root_array[index]
    
def select_drive_letter():
    global drive_letter, destination_root, playlist_root
    prompt = "What drive letter does it have now (just the letter): "
        
    index = 0    
    while True:    
      ans = raw_input(prompt)
      if ans == "":
        continue
      else:
        break
    
    playlists = playlists_array[index]
    destination_root = '%s:\%s' % (ans, destination_root)
    playlist_root = '%s:\%s' % (ans, playlist_root)
    
def confirm(prompt=None, resp=False):
    """prompts for yes or no response from the user. Returns True for yes and
    False for no.

    'resp' should be set to the default value assumed by the caller when
    user simply types ENTER.

    >>> confirm(prompt='Create Directory?', resp=True)
    Create Directory? [y]|n: 
    True
    >>> confirm(prompt='Create Directory?', resp=False)
    Create Directory? [n]|y: 
    False
    >>> confirm(prompt='Create Directory?', resp=False)
    Create Directory? [n]|y: y
    True

    """
    
    if prompt is None:
        prompt = 'Confirm'

    if resp:
        prompt = '%s [%s]|%s: ' % (prompt, 'y', 'n')
    else:
        prompt = '%s [%s]|%s: ' % (prompt, 'n', 'y')
        
    while True:
        ans = raw_input(prompt)
        if not ans:
            return resp
        if ans not in ['y', 'Y', 'n', 'N']:
            print 'please enter y or n.'
            continue
        if ans == 'y' or ans == 'Y':
            return True
        if ans == 'n' or ans == 'N':
            return False

if __name__ == "__main__":  
    main()