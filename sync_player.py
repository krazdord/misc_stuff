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
#

########################
#### SETTINGS BELOW ####
########################

# this is an array of foobar2000 playlist names you want synched
#
playlists = [ "car", "test" ]  

# this is the path to your android/mp3 player music folder once mounted.
# the converted files will be placed here.
#
destination_root = r"F:\Test_Music"

# this is the path to your android/mp3 player playlist folder. m3u files will
# place here.
#
playlist_root = r"F:\Test_Music\Playlists"

# this is your target conversion format.
#
destination_ext = ".mp3"

# this is how many path levels of your source dir to ignore
# ie - if you put music in X:\Music\stuff in foobar,
# this should be 2 to make a folder called stuff under
# destination_root defined above
# 
path_ignore_depth = 2

# change these paths to reflect where your converters are installed
# 
#ffmpeg_path = r"C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe"
#lame_path = r"C:\Program Files (x86)\LAME\lame.exe"
#flac_path = r"C:\Program Files (x86)\FLAC\flac.exe"
#requires lame, sed, flac, flacmeta on the path

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
from urlparse import urlparse
from logging import StreamHandler, Formatter, DEBUG, INFO, getLogger
from shutil import copyfile

log = getLogger("Foobar.MP3PlayerSync")
log.setLevel(DEBUG)
lh = StreamHandler()
lh.setFormatter(Formatter("%(levelname)s|%(asctime)s|%(filename)s:%(lineno)d|%(message)s"))
log.addHandler(lh)

log.info("Connecting to foobar2000 COM server...")

prog_id = "Foobar2000.Application.0.7"

fb2k = win32com.client.Dispatch(prog_id)
all_files = []
files_copied = 0
files_transcoded = 0
files_skipped = 0
files_deleted = 0

def main():
    global all_files
    global files_deleted
    all_files = scan_dir(destination_root, [destination_ext])
    print "Found " + str(len(all_files)) + " files on the device."
    
    fb2k_playlists = [ i for i in fb2k.Playlists if i.Name in playlists ]
    if fb2k_playlists:
        for pl in fb2k_playlists:
            sync_playlist(pl)
    log.info("After sync %i files not matched to a playlist." % len(all_files))
    if len(all_files) > 0:
      confirm("Delete these files?", True)
      files_deleted = len(all_files)
      for each_file in all_files:
        log.info("Removing: %s" % each_file.encode(sys.getfilesystemencoding()))
        remove(each_file.encode(sys.getfilesystemencoding()))
        
    log.info("Completed Sync!")
    log.info("Copied %i files, transcoded %i files, skipped %i files, deleted %i files" % (files_copied, files_transcoded, files_skipped, files_deleted))
    
    
def scan_dir(path, exts):
    """
    path    -    where to begin folder scan
    """
    selected_files = []
    
    for root, dirs, files in walk(path): 
        selected_files += select_files(root, files, exts)
            
    return selected_files

    
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
            selected_files.append(unicode(full_path.decode(sys.getfilesystemencoding())))
          #else:
           # print "Extensions didn't match (" + ext + " f:" + file_ext + "), skipping: " + full_path

    return selected_files   
    

def sync_playlist(sync_playlist):
    log.info("Syncing playlist '%s'..." % sync_playlist.Name)
    tracks = sync_playlist.GetTracks()
    
    m3u_lines = ["#EXTM3U"]
    for t in tracks: 
        m3u_lines.append(t.FormatTitle("#EXTINF:%length_seconds%, %artist% - %title%"))
        source_path = urlparse(t.Path).netloc
        dest_path = sync_file(source_path)
        m3u_lines.append(dest_path)
        #print "adding line to m3u: " + dest_path.encode(sys.getfilesystemencoding()) 
        idx = 0
        try:     
          idx = all_files.index(unicode(dest_path))
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
   
    parts_new_path = [destination_root]        
    parts_new_path.extend(parts[path_ignore_depth:length-1])
    destination_folder = sep.join(parts_new_path)
    parts_new_path.append(filename + destination_ext)     
    destination_path = sep.join(parts_new_path)
    
    if not exists(destination_folder):
        log.debug("Creating folder: '%s'..." % destination_folder)
        makedirs(destination_folder)
    
    if not exists(destination_path):
      if (destination_ext.lower() != ext.lower()):
        convert_file(source_path, destination_path)
        files_transcoded += 1     
      else:
        log.info("Copying: '%s' -> '%s'" % (source_path, destination_path))
        copyfile(source_path, destination_path)
        files_copied += 1
    else:
      files_skipped += 1
        
    return unicode(destination_path)
    
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
    
    
def convert_file(input_file, output_file):

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

    #TODO: call encode once? im lazy, eh
    log.debug("Converter command line:\n%s" % command.encode(sys.getfilesystemencoding()))
    try:
        retcode = call(command.encode(sys.getfilesystemencoding()), shell=True)        
    except OSError, e:
        log.critical("Converter execution failed: '%s'", e.strerror)
        
 
def create_m3u(playlist_name, m3u_lines):
    if not exists(playlist_root):
        log.info("Creating folder: '%s'..." % playlist_root)
        makedirs(playlist_root)
        
    m3u_path = "%s\\%s.m3u" % (playlist_root, playlist_name)
    log.info("Creating m3u playlist: '%s'..." % m3u_path)    
    f = codecs.open(m3u_path, 'w', 'mbcs')
    f.write("\n".join(m3u_lines))
    f.close()
    
    
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