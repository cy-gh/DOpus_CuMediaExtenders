# CuMediaExtenders
CuMediaExtenders is a script-addin for video and audio files for the marvellous [Directory Opus file manager](https://www.gpsoft.com.au/), using MediaInfo &amp; NTFS ADS as its backend.
It complements & enhances DOpus built-in fields, instead of completely replacing them.

As of v0.9/20210113 you get:
* 15 commands incl. custom icons
* 41 columns/script fields
* DCF menu with all available commands
* Sample JSON file to customize column labels (requires Script or DOpus restart)
* PSD template used to create the icons

It has extensive configuration options with descriptions and include samples in config section 'Reference Only'.

Column labels are customizable with a JSON file, a sample file is in distro (extract OSP if not visible).

Video & Audio Codec, # of Audio Channels, Duration groups and Resolution fields are customizable without modifying the script.

'Dirty' files, i.e. files which are changed since writing to ADS, can be detected.

Seen files are cached in DOpus memory, to speed up reading process (deactivatable).

Progress dialog with Pause/Resume & Abort is supported.

4 additional buttons are supplied which can toggle on/off field groups, which are customizable.



![./Screenshots/00-Showcase.png](./Screenshots/00-Showcase.png)



Read the DISCLAIMER! If you blame me for anything nonetheless, you will be  readily  ignored.

December 2020/January 2021

GitHub: https://github.com/cy-gh/DOpus_CuMediaExtenders



## Quick Start

  * Download [MediaInfo CLI/Portable](https://mediaarea.net/en/MediaInfo/Download/Windows) (not tested with Setup version)
  * Open script config and set path to MediaInfo.exe
  * Add UPDATE and other optional commands to any toolbar/menu
  * Add columns (you can use supplied TOGGLE commands as well)
  * Process some files with UPDATE
  * ...and voila, happy organizing!

Although not required, this setting is HIGHLY RECOMMENDED:

  **File Operations->Copy Attributes->Copy all NTFS streams: TRUE**

This ensures that if you copy a file the Metadata ADS is copied with it.



## License

CuMediaExtenders v0.9 - © 2021 cuneytyilmaz.com

[Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)](https://creativecommons.org/licenses/by-sa/4.0/)

[MediaInfo](https://mediaarea.net/en/MediaInfo) & MediaInfo icon: Copyright (c) 2002-2021 MediaArea.net SARL

sprintf: https://hexmen.com/blog/2007/03/14/printf-sprintf/

Android Material icons Copyright (c) Google https://material.io/resources/icons/



## Supported Formats (as of v0.9)

Containers && Codecs I have tested so far:

​	MKV, MP4, AVI, FLV, WEBM, 3GP,

​	M4A, M4B, MKA, MP3, MP2, MP1, FLAC, AC3, AAC, DTS, TrueHD (Dolby Atmos),

​	Wave, Wave64, ALAC, TAK, TTA, DSD, Ogg Vorbis, AIFF, AMR, WavPack, WMA Lossy & Lossless, MusePack

....and any file as long as MediaInfo reports at least a video or audio track.



This means:

1. To process multiple files you can select whatever you want, not just video or audio files, and any file which does not have at least 1 video or audio track will be skipped and no ADS data will be stored.
   However, you might want to select less files to avoid unnecessary probing into the non-multimedia files.
2. You can customize the output for any format I forgot or new formats recognized by MediaInfo. If format and/or codec information is shown, the displayed string can be customized arbitrarily WITHOUT reprocessing files.

A full list can be found @ https://mediaarea.net/en/MediaInfo/Support/Formats



Not processed by definition:

​	Directories, reparse points, junctions, symlinks



Some file types are better supported than others, e.g. RG can be saved in WAVs (by Foobar2000) but not recognized by MediaInfo.

## Available Fields

Columns are by default prefixed with ME (Media Extensions); labels can be customized via external config file.

Note although some of these fields also exist in DOpus, they are not available for all container & codec types

| Field (not IDs) | Desc                                                         |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|HasADS/IsAvailable 		| File has Metadata ADS (calculated separately, not Multicol) [++]|
|NeedsUpdate/Dirty 		| File has been changed since Metadata ADS has been written [++]|
|VCodec           		| Video codec (only 1st video stream) [+]|
|ACodec           		| Audio codec (only 1st audio stream) [+]|
|TBitrate       			| Total bitrate [++]|
|VBitrate           		| Video bitrate (only 1st video stream) [+]|
|ABitrate           		| Audio bitrate (only 1st audio stream) [+]|
|VCount           		| Number of video streams [++]|
|ACount           		| Number of audio streams [++]|
|TCount            		| Number of text (subtitle) streams [++]|
|OCount           		| Number of other (chapters, menus...) streams [++]|
|ABitrate Mode    		| VBR/CBR if available (ABR is principally recognized as VBR) [++]|
|TDuration       		| Total container duration [+]/[++]|
|VDuration    			| Video duration [+]/[++]|
|Auration          		| Audio duration [+]/[++]|
|Duration (Combo) 		| Combined duration, only a single one is displayed unless any of 3 differs from each other (5s tolerance), i.e. detects mismatching streams [++]|
|Multi-Audio      		| File has multiple audio streams (yes/no) [++]|
|Audio Channels   		| Number of channels in 1st audio track, e.g. 2.0, 5.1, 7.1  [+]|
|ReplayGain       		| File has ReplayGain (RG) info stored (some format support RG but MediaInfo does not parse it) [++]|
|Audio Compression Mode 	| Lossy/Lossless [++]|
|AspectRatio (Raw) 		| Video aspect ratio (AR), i.e. width:height, e.g. 1.333, 2.35, etc. [++]|
|AspectRatio (Disp) 		| Video display AR if the format supports it and file has a different AR than its raw AR [++]|
|AspectRatio (Combined) 	| Combined video AR, only a single AR is displayed, if its display AR differs from raw AR, it's shown in parentheses [++]|
|Gross byterate     		| Gross average KB, kilobytes per second, for the file, simply filesize:duration [++]|
|Bits per Pixel     		| Video average bits per pixel, i.e. video bitrate/(width*height); only 1st video stream [++]|
|Dimensions       		| Video dimensions [+]|
|Resolution       		| 240p, 360p, 480p, 720p, 1080p, 4K, 8K - Customizable, can append '(Vertical)' for vertical videos [++]|
|Frame Rate       		| Video frame rate [+]|
|Subtitle Language 		| Subtitle language if available (only 1st text stream) [++]|
|Encoded Library   		| Library used to encode the video if available, incl. technical info [++]|
|Encoded Lib Name 		| Library used to encode the video if available no technical info  [++]|
|VCodec ID       		| Video Codec ID (raw), if available [++]|
|ACodec ID       		| Audio Codec ID (raw), if available [++]|
|AFormat Version  		| Audio Format Version, if available [++]|
|AFormat Profile 		| Audio Format Profile, if available [++]|
|Encoder App      		| Container Encoder App, if available [++]|
|ADS (Formatted) 		| Formatted ADS data (always as JSON), suitable to show in InfoTips [++]|
|ADS (Raw)       		| Unformatted ADS data (always as JSON), only for sake of completeness, not very suitable as a column or InfoTip [++]|
|Helpers            		| 3 helper columns for the container, video & audio tracks, which you can use to adjust display Codec names [++]|

[+]   Recognizes more formats and/or is more accurate than DOpus

[++]  No DOpus counterpart



Unlike DOpus, output of following columns are user-customizable (editing source code is also another option):

* VCodec
* ACodec
* Audio Channels
* Resolution

Although there are no fields for the following, they are stored and available in ADS (see section Adding/Adjusting Fields):

* Overall Bitrate Mode
* Additional & Extra info fields for Container, Video, Audio & Text
    These container- and codec-specific fields can contain a lot information e.g. ReplayGain, used encoding settings, date-time of encoding... which can be used show new fields/columns or adjust existing ones.
* Video Stream Size
* Audio Stream Size
* Audio Format Settings Mode
* Audio Sampling Rate



Note some of the fields (e.g. sampling rate) above are already supported by DOpus well close to 100%, albeit not for all containers & codecs.

Any other info should be used from DOpus builtin fields.



## Available Actions

* Update Metadata: Create/Update Metadata ADS for selected file
* Delete Metadata: Delete attached Metadata ADS of selected file
* Clear Cache: Clear in-memory DOpus cache for Metadata read from ADS (see FEATURES)
* Copy Metadata: Copy Metadata ADS of selected files to clipboard
* Dump Metadata: Dump Metadata ADS of selected files to DOpus Output window
* Copy MediaInfo: Run MediaInfo anew and copy its output to clipboard
* Dump MediaInfo: Run MediaInfo anew and dump its output to DOpus Output window
* Estimate Bitrate: Calculate bitrate using a list of target 'bitrate/pixel' values
* Toggle Essential: Toggle a user-configurable list of 'essential' columns, can toggle DOpus columns
* Toggle Optional: Toggle a user-configurable list of 'optional' columns, can toggle DOpus columns
* Toggle Other: Toggle a user-configurable list of 'other' columns, can toggle DOpus columns
* Toggle Verbose: Toggle a user-configurable list of 'verbose' columns, can toggle DOpus columns
* Validate Config: Validate current user configuration



A preconfigured menu button with all available actions is in distro (extract OSP if not visible); see screenshots below.

## Customization

If you want to adjust the VCodec & ACodec output, see options:

* CODEC_APPEND_ADDINFO
* CODEC_USE_SHORT_VARIANT
* RESOLUTION_APPEND_VERTICAL
* LOOKUP_CODECS - as a LAST resort!

LOOKUP_CODECS can be used to show identical info to FourCC if you want, but I tell you FourCC is pretty useless, it's not standardized nor filled/respected/recognized by many encoders/muxers. 

Suggested: turn on one or more of the 'HELPER' columns (Toggle Verbose action) to see what info you can use to identify a certain codec.



If you want to rename the labels, you need to create a JSON file in the script directory.



First check the formatted Metadata ADS, e.g. via Tooltip or Dump/Copy for a selected file.

Some fields are already stored in ADS, but have no dedicated columns, and you'd need to create only a new column.

Chances are very high that the information you need is already there (except FourCC, since it is a placebo BS field).



If you want to add new buttons or columns based on available info, see OnInit() method

and use one of the existing commands & fields as template; note few fields e.g. 'Available' does not use MultiCol!



If the info you want is not in the ADS, copy/dump the MediaInfo output for the file.

If MediaInfo does recognize the field you want, check OnME_Update() and add the field. **HOWEVER,** adding new fields means adding new fields to the structure stored in ADS and that means, you will need to update previously processed files with the ADS before you can use new fields!

Please keep that in mind.



### Feature Requests:

Requesting new fields/user option is ok, with some conditions:

* No obscure fields, like encoder 'Format_Settings_Matrix', should be usable by many

* My time allows it

* You must supply a sample file and MediaInfo output of the sample, where the field(s) you want are marked clearly



## Features

### Nearly constant scanning speed

File scanning speed does NOT depend on the individual file sizes since most containers store metadata at the front of the files and this is what MediaInfo parses. This means, the overall scanning duration depends more or less on the total number of files, not their sizes. The scanning speed depends on your disk's read speed of course but since SDD & NVMEs ubiquitous these days the speed is pretty acceptable.

### Dirty detection:

If a file has been changed, e.g. overwritten by a program or its file-specific metadata (ID3, keywords, etc.) have been updated, this is automatically recognized and can be used to update its Metadata ADS.

#### Progress window

"Pause/Resume" & "Abort" actions are supported fully, Skip seemed unnecessary to me since scanning speed per file is pretty high.

### Caching:

This is an experimental feature and does not bring as much performance gain as you might expect from in-memory caching.

Can be disabled via option.

Disabling the cache and/or clearing it does not seem to make a huge impact; YMMV. Read on if you're curious.

Any seen file, as long as the columns are visible, are cached in DOpus memory, ~1k of data per file. If the file path, name & last modification time have not been changed, the cached info is used.

The Metadata info is principally processed using POJO/JSON, and my initial concern was reading JSON strings, converting them to POJO and back would be very time-consuming has been removed by another factor.

DOpus seems to have a minimum overhead for script columns/fields ; i.e. even if you program the simplest fields/columns, updating the columns in a big directory still takes some time, although DOpus-internal fields are shown instantly.

However, after much testing, it is safe to say, caching does not speed up the refresh speed of columns, since DOpus always has an overhead to display script fields/columns, even if they show a simple 'X' without calculating any information at all.

The caching only reduces disk accesses, since NTFS ADS streams are practically separate files.

The cached info is 'stringified JSONs', i.e. POJOs (Plain Old Javascript Objects) converted to readable format. This information is read and parsed (converted back to an in-memory POJO) whenever needed. Even if this takes some computing time, DOpus does not allow you to keep POJOs in memory for too long, these are removed from memory as soon as DOpus show 'Script Completed' in output window, even if you use Script.Vars object. Script.Vars can store simple types (integer, string, etc.) or DOpus objects, e.g. Vector & Map but not POJOs. Since it is too much hassle for me to convert all JS object accesses from object.field to DOPus Map() notation, I'm not planning to rewrite it and cache directly the parsed JSONs. Considering script-addins do not run faster than a certain speed, a rewrite would not bring much speed increase anyway. The only possible solution seems to be developing a DOpus plugin, which I'm not planning at the moment.

Also added a parameter cache_max_items to limit the number of items in the cache based on timestamp of the item added to the cache but iterating over hash arrays based on an attribute other than the hash's key field turned out to be an unnecessary hassle too much pain for very little gain, so such a feature won't be implemented (again), just use the CLEAR CACHE command

### Multiple Video Tracks:

  You might wonder why "only 1st video stream" is mentioned elsewhere.

  The wonderful Matroska format allows to mux multiple video streams.

  But unlike files with multiple audio streams, no media player I know of can play such files in entirety,

  However, for the sake of completeness & OCD I have included them in the script.

  

## Features & Screenshots

This is what you should see when you install the script

![DOpus Scripts preferences](./Screenshots/01.png)





Script Preferences![02](./Screenshots/02.png)





Pay also attention to the group headers & descriptions, they will save you some head scratching. The last section with REF_xxx fields is for reference only, changing these will have no effect.

Some changes can be directly used without an UPDATE, i.e. rescanning the files and updating the ADS data. Only few, luckily they don't change much, such as which types of codecs are lossless/lossy or VBR need an ADS UPDATE, in case you use formats which I have not tested with yet.

For example, for the setting LOOKUP_CODECS, there's a corresponding REF_LOOKUP_CODECS, which includes many comments.

![03](./Screenshots/03.png)





Sample setting for channel count lookup. This must be valid JSON as the description states. With this you can list files as 2.0, 5.1, 7.1 etc. or mono, stereo, etc. The \u2260 is the Unicode ≠ (not equal sign).

![04](./Screenshots/04.png)





Duration groups which are used when you group your files by duration.

![05](./Screenshots/05.png)





Resolution lookup, which are used both for the Resolution column and the grouping; the ""(Vertical)" suffix can be separately added via the option RESOLUTION_APPEND_VERTICAL.

![06](./Screenshots/06.png)





In distro file you will find some buttons which can toggle 4 groups of fields on and off. This is the "Essential Fields" list, which I personally find very useful in addition to DOpus builtin fields. You can also mix it up with DOpus fields like mp3songlength, and toggle them whenever you need them.

The next option TOGGLEABLE_FIELDS_ESSENTIAL_AFTER is any field name **after** which these fields are shown. It can be any field, incl. this script's fields. If empty, fields are added to/removed from the last column.

![07](./Screenshots/07.png)





Sample reference variable for external config file. These are automatically converted strings from script source code (via a very ugly but very effective hack). Use these REF fields to read unstripped objects; the real variables they correspond to must be valid JSON; remember what I said above about descriptions, always read them.

More to this in the screenshots below.

![08](./Screenshots/08.png)





Some fields more suitable to use in Infotips or to dump into Output Window than others.![09](./Screenshots/09.png)





A sample pool of files I used to test the columns. This is the "Essential Fields" in the middle and DOpus fields on the right. Study it and compare them as much as you need.

![10](./Screenshots/10.png)



"Optional Fields"![11](./Screenshots/11.png)



"Other fields" - Can be used to tap into already existing ADS data.

![12](./Screenshots/12.png)



"Verbose fields" - You wouldn't want to use them as columns. The Formatted version is more suitable for InfoTips and raw one is for copy-paste into an editor.

![13](./Screenshots/13.png)





Supplied commands as of v0.9.

![14](./Screenshots/14.png)



A ready-to-use menu as DCF file to get you started.

![15](./Screenshots/15.png)





There's no accounting for taste. The column names I personally find useful might not suit your needs and likes. So I made them customizable. Since DOpus reads column headers only during script initialization, there is no way to make them user-configurable... until you create a JSON file as described below in the script folder. See screenshots further below how it is used.

![16](./Screenshots/16.png)



The about screen

![17](./Screenshots/17.png)



The "codec" fields can be quite overwhelming. Before you start "but FourCC..." shush shush... No! It's a bullshit field, which is neither standardized nor respected by many editors, encoder and muxers. In the past, some encoders even misused it deliberately to maintain compatibility with existing codecs, *cough DivX 3, MP42, XviD, DivX 5... cough*.

Since video encoding is a world of its own and each codec has a gazillion options and combinations, you have the choice to customize the magical "codec" info as you please. 

Below is the "Additional Info" setting which affects only a handful of codecs like AAC, DTS, Atmos. This info comes straight from MediaInfo, except Opus files, which the script fills separately for you. Note how "AAC (LC SBR)" becomes "AAC" and so on, but not MP3.

![18](./Screenshots/18.png)



You might be confused why MP3 (v1) was identical on both sides in the screenshot above, because this is technically not from "Additional Info" field, but you can still use the "short versions" as seen below.

![19](./Screenshots/19.png)





Where do these "short versions" of codecs come from? From this list. It's not easy to use, but quite powerful if you take the time to look into the existing ADS data (and there are 2 buttons for that) and adjust the LOOKUP_CODECS (must be valid JSON).

![20](./Screenshots/20.png)



## Future

### Extra Music fields?

Since Foobar2000 & MP3Tag both are stellar programs, I see little need at the moment. Probably a 'Cover' field, which is also used in MKVs and I plan to use it.

### Images?

DOpus already does a terrific job with images, I see little need at the moment, maybe a DPI field... but DOpus probably has it somewhere I haven't seen yet :D

### Community Feedback

Although I have tested the script so far with many codecs I daily use or simply know of the world of multimedia is a jungle and some files, container/codec combinations are surely missed.

Speed, how fast these columns are parsed/displayed, is another topic where I have no idea how it'd work for others. It's pretty fast on my machine, but I do have NVMEs & 12-core machine, so... no idea!

## Version History

v0.5:

Initial release

...

many unreleased versions, see source if you're really that interested... I thought so.

...
