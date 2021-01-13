# DOpus_CuMediaExtenders
Video &amp; audio extended script fields &amp; columns for Directory Opus file manager, with MediaInfo &amp; NTFS ADS backend

CuMediaExtenders is a script-addin for DOpus for video and audio files, using MediaInfo and NTFS ADS as its backend.
It complements & enhances DOpus built-in fields, instead of completely replacing them.

As of v0.9/20210113 you get:
* 15 commands incl. custom icons
* 41 columns/script fields
* DCF menu with all available commands
* Sample JSON file to customize column labels (requires Script or DOpus restart)
* PSD template used to create the icons

It has extensive configuration options with descriptions
and include samples in config section 'Reference Only'.

Column labels are customizable with a JSON file, a sample file is in distro (extract OSP if not visible).

Video & Audio Codec, # of Audio Channels, Duration groups and Resolution fields are customizable without modifying the script.

'Dirty' files, i.e. files which are changed since writing to ADS, can be detected.

Seen files are cached in DOpus memory, to speed up reading process (deactivatable).

Progress dialog with Pause/Resume & Abort is supported.

4 additional buttons are supplied which can toggle on/off field groups, which are customizable.

I suggest reading all sections; chances are your potential question has already been answered.

Read the DISCLAIMER! If you blame me for anything nonetheless, you will be  readily  ignored.

Quick Start:
  * Download MediaInfo CLI/Portable (not tested with Setup version)
  * Open script config and set path to MediaInfo.exe
  * Add UPDATE and other optional commands to any toolbar/menu
  * Add columns (you can use supplied TOGGLE commands as well)
  * Process some files with UPDATE
  * ...and voila, happy organizing!

Although not required, this setting is HIGHLY RECOMMENDED:
  **File Operations->Copy Attributes->Copy all NTFS streams: TRUE**
  
This ensures that if you copy a file the Metadata ADS is copied with it.



CuMediaExtenders v0.9 - © 2021 cuneytyilmaz.com

Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)

MediaInfo & MediaInfo icon: Copyright (c) 2002-2021 MediaArea.net SARL

sprintf: https://hexmen.com/blog/2007/03/14/printf-sprintf/

Android Material icons Copyright (c) Google https://material.io/resources/icons/
