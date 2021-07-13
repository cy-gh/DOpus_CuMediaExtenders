// @ts-nocheck
/* eslint-disable quotes */
/* eslint-disable no-inner-declarations */
/* global ActiveXObject Enumerator DOpus Script */
/* eslint indent: [2, 4, {"SwitchCase": 1}] */
///<reference path="./_DOpusDefinitions.d.ts" />

// GLOBAL objects
{
    /**
	 * do not touch
	 */
    var Global = { };
    Global.SCRIPT_NAME        = 'CuMediaExtenders'; // WARNING: if you change this after initial use you have to reconfigure your columns, infotips, rename scripts...
    Global.SCRIPT_NAME_SHORT  = 'CME';
    Global.SCRIPT_VERSION     = 'v0.93';
    Global.SCRIPT_COPYRIGHT   = '© 2021 cuneytyilmaz.com';
    Global.SCRIPT_URL         = 'https://github.com/cy-gh/DOpus_CuMediaExtenders/';
    Global.SCRIPT_DESC        = 'Extended fields for multimedia files (movie & audio) with the help of MediaInfo & NTFS ADS';
    Global.SCRIPT_MIN_VERSION = '12.23';
    Global.SCRIPT_DATE        = '20210713';
    Global.SCRIPT_GROUP       = 'cuneytyilmaz.com';
    Global.SCRIPT_PREFIX      = 'MExt';				// prefix for field checks, log outputs, progress windows, etc. - do not touch
    Global.SCRIPT_LICENSE     = 'Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)';

    var util = {};
    util.cmdGlobal            = DOpus.Create.Command;
    util.stGlobal             = DOpus.Create.StringTools;
    util.shellGlobal          = new ActiveXObject('WScript.shell');
    util.sv                   = Script.vars;

    // used by ReadFile() & SaveFile()
    var TEXT_ENCODING = { 'utf8': 1, 'utf16': 2 };
}


String.prototype.trim = function () {
    return this.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
}
String.prototype.normalizeLeadingWhiteSpace = function () {
    return this
        .replace(/^\t\t\t|^ {12}/mg, '§§')
        .replace(/^\t\t|^ {8}/mg, '§')
        .replace(/^\t|^ {4}/mg, '')
        .replace(/§/mg, '  ')
        .replace(/\n/g, '\r\n')
        .trim();
};
String.prototype.substituteVars = function () {
    return this.replace(/\${([^}]+)}/g, function (match, p1) {
        return typeof eval(p1) !== 'undefined'
            ? eval(p1)
            : 'undefined'
        ;
    });
};
// HEREDOC objects
var HEREDOCS = {
    // ugly but functional; JScript under Windows lacks many cool features of JavaScript
    'Read First': function(){/*
		${Global.SCRIPT_NAME} is a script-addin for DOpus for video and audio files,
		using MediaInfo and NTFS ADS as its backend.
		It complements & enhances DOpus built-in fields, instead of completely replacing them.

		As of ${Global.SCRIPT_VERSION}/${Global.SCRIPT_DATE} you get:
		* 15 commands incl. custom icons
		* 40+ columns/script fields
		* DCF menu with all available commands
		* Sample JSON file to customize column labels (requires Script or DOpus restart)
		* PSD template used to create the icons

		It has extensive configuration options with descriptions
		and include samples in config section 'Reference Only'.

		Column labels are customizable with a JSON file, a sample file is in distro (extract OSP if not visible).

		Video & Audio Codec, # of Audio Channels, Duration groups and Resolution fields
		are customizable without modifying the script.

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
			File Operations->Copy Attributes->Copy all NTFS streams: TRUE
		This ensures that if you copy a file the Metadata ADS is copied with it.

		${Global.SCRIPT_NAME} ${Global.SCRIPT_VERSION} - ${Global.SCRIPT_COPYRIGHT}
		${Global.SCRIPT_LICENSE}
		MediaInfo & MediaInfo icon: Copyright (c) 2002-2021 MediaArea.net SARL
		sprintf: https://hexmen.com/blog/2007/03/14/printf-sprintf/
		Android Material icons Copyright (c) Google https://material.io/resources/icons/
	*/}.toString().substituteVars(),


    'What Is This': function(){/*
		${Global.SCRIPT_NAME} turns DOpus to a full-fledged video/audio manager.

		It supplies extended info columns & script fields, and script commands for multimedia files
		which you can use everywhere these are available: Listers, InfoTips, Tile view, Rename Dialog...
		...as long as the files are processed before.

		'Processing' a file (UPDATE command) means storing crucial information about a file in ADS
		which can be later queried faster, and combined & formatted for display.

		It recognizes many container formats and codecs which DOpus supports only partially or not at all.

		The values are supplied by the external program MediaInfo (CLI version)
		which can be freely downloaded.

		The values are NOT stored in the files themselves, as in MP3 ID3, Matroska, APE tags, etc.
		but attached to the files.

		This script uses ADS and metadata, and occasionally stream, interchangably.
		Note: Streams are not to be confused with video or audio streams in the files.

		"Attached"?
		Windows NTFS supports the so-called Alternate Data Streams (ADS).
		These are practically separate 'quasi-files' attached to their parent files.
		0 or more ADS streams can be attached to a file.
		These streams can be listed with "DIR /R" (CMD) or "DIR /:" (JPSoft TCC).

		You might know ADS from files downloaded from Internet.
		Whenever you download a file with IE/Edge, it appends a 'ZoneInformation' ADS to the file
		and if you execute it Windows asks you before if you trust the file source and want to execute it.

		More information about ADS can be found via a simple search "NTFS ADS", e.g.

		https://en.wikipedia.org/wiki/NTFS#Alternate_data_streams_(ADS)
		https://www.2brightsparks.com/resources/articles/ntfs-alternate-data-stream-ads.html
		https://stealthbits.com/blog/ntfs-file-streams/
		https://hshrzd.wordpress.com/2016/03/19/introduction-to-ads-alternate-data-streams/

		Are ADS safe?
		Normally when you edit the file content, the ADS content is not changed and vice versa,
		that is, it is normally safe to edit the file and its stream(s) separately.
		When you delete the parent file, its ADS are automatically deleted.

		When you change the metadata stream's content, its parent file's modification timestamp may be updated,
		however, this script allows you to keep the original timestamp (highly recommended).

		ADS can be used for good and bad. Some malware can hide their code in ADS,
		however, most modern antivirus software recognize this already.

		On any non-NTFS file system this script will not work.
	*/}.toString().substituteVars(),


    'Supported Formats': function(){/*
		Containers && Codecs I have tested so far:

		MKV, MP4, AVI, FLV, WEBM, 3GP,
		M4A, M4B, MKA, MP3, MP2, MP1, FLAC, AC3, AAC, DTS, TrueHD (Dolby Atmos),
		Wave, Wave64, ALAC, TAK, TTA, DSD, Ogg Vorbis, AIFF, AMR, WavPack, WMA Lossy & Lossless, MusePack

		....and any file as long as MediaInfo reports at least a video or audio track.

		This means:
		1.	To process multiple files you can select whatever you want, not just video or audio files,
		   	and any file which does not have at least 1 video or audio track will be skipped and no ADS data will be stored.
		   	However, you might want to select less files to avoid unnecessary probing into the non-multimedia files.
		2.	You can customize the output for any format I forgot or new formats recognized by MediaInfo.
		   	If format and/or codec information is shown, the displayed string can be customized arbitrarily WITHOUT reprocessing files.

		A full list can be found @ https://mediaarea.net/en/MediaInfo/Support/Formats

		Not processed by definition:
			Directories, reparse points, junctions, symlinks

		Some file types are better supported than others,
		e.g. RG can be saved in WAVs (by Foobar2000) but not recognized by MediaInfo.
	*/}.toString().substituteVars(),


    'Available Fields': function(){/*
		Columns are by default prefixed with ME (Media Extensions); labels can be customized via external config file.

		Note although some of these fields also exist in DOpus, they are not available for all container & codec types

		HasADS/IsAvailable:		File has Metadata ADS (calculated separately, not Multicol) [++]
		NeedsUpdate/Dirty: 		File has been changed since Metadata ADS has been written [++]
		VCodec:          		Video codec (only 1st video stream) [+]
		ACodec:          		Audio codec (only 1st audio stream) [+]
		TBitrate:      			Total bitrate [++]
		VBitrate:          		Video bitrate (only 1st video stream) [+]
		ABitrate:          		Audio bitrate (only 1st audio stream) [+]
		VCount:          		Number of video streams [++]
		ACount:          		Number of audio streams [++]
		TCount:           		Number of text (subtitle) streams [++]
		OCount:          		Number of other (chapters, menus...) streams [++]
		ABitrate Mode:   		VBR/CBR if available (ABR is principally recognized as VBR) [++]
		TDuration:      		Total container duration [+]/[++]
		VDuration:   			Video duration [+]/[++]
		Auration:         		Audio duration [+]/[++]
		Duration (Combo):		Combined duration, only a single one is displayed unless any of 3 differs from each other (5s tolerance), i.e. detects mismatching streams [++]
		Multi-Audio:     		File has multiple audio streams (yes/no) [++]
		Audio Language:     	Audio language if set (only few containers support this) [++]
		Audio Channels:  		Number of channels in 1st audio track, e.g. 2.0, 5.1, 7.1  [+]
		ReplayGain:      		File has ReplayGain (RG) info stored (some format support RG but MediaInfo does not parse it) [++]
		Audio Compression Mode:	Lossy/Lossless [++]
		AspectRatio (Raw):		Video aspect ratio (AR), i.e. width:height, e.g. 1.333, 2.35, etc. [++]
		AspectRatio (Disp):		Video display AR if the format supports it and file has a different AR than its raw AR [++]
		AspectRatio (Combined):	Combined video AR, only a single AR is displayed, if its display AR differs from raw AR, it's shown in parentheses [++]
		Gross byterate:    		Gross average KB, kilobytes per second, for the file, simply filesize:duration [++]
		Bits per Pixel:    		Video average bits per pixel, i.e. video bitrate/(width*height); only 1st video stream [++]
		Dimensions:      		Video dimensions [+]
		Resolution:      		240p, 360p, 480p, 720p, 1080p, 4K, 8K - Customizable, can append '(Vertical)' for vertical videos [++]
		Frame Rate:      		Video frame rate [+]
		Frame Count:      		Video frame count [++]
		Subtitle Language:		Subtitle language if available (only 1st text stream) [++]
		Encoded Library:  		Library used to encode the video if available, incl. technical info [++]
		Encoded Lib Name:		Library used to encode the video if available no technical info  [++]
		VCodec ID:      		Video Codec ID (raw), if available [++]
		ACodec ID:      		Audio Codec ID (raw), if available [++]
		AFormat Version: 		Audio Format Version, if available [++]
		AFormat Profile:		Audio Format Profile, if available [++]
		Encoder App      		Container Encoder App, if available [++]
		ADS (Formatted):		Formatted ADS data (always as JSON), suitable to show in InfoTips [++]
		ADS (Raw):      		Unformatted ADS data (always as JSON), only for sake of completeness, not very suitable as a column or InfoTip [++]
		Helpers:           		3 helper columns for the container, video & audio tracks, which you can use to adjust display Codec names [++]

		[+] 	Recognizes more formats and/or is more accurate than DOpus
		[++]	No DOpus counterpart

		Unlike DOpus, output of following columns are user-customizable (editing source code is also another option):
			VCodec, ACodec, Audio Channels, Resolution

		Although there are no fields for the following, they are stored and available in ADS (see section Adding/Adjusting Fields)

			Overall Bitrate Mode
			Additional & Extra info fields for Container, Video, Audio & Text
				These container- and codec-specific fields can contain a lot information
				e.g. ReplayGain, used encoding settings, date-time of encoding...
				which can be used show new fields/columns or adjust existing ones.
			Video Stream Size
			Audio Stream Size
			Audio Format Settings Mode
			Audio Sampling Rate

		Note some of the fields (e.g. sampling rate) above are already supported by DOpus well close to 100%, albeit not for all containers & codecs.
		Any other info should be used from DOpus builtin fields.
	*/}.toString().substituteVars(),


    'Available Actions': function(){/*
		Update Metadata:		Create/Update Metadata ADS for selected file
		Delete Metadata:		Delete attached Metadata ADS of selected file
		Clear Cache:    		Clear in-memory DOpus cache for Metadata read from ADS (see FEATURES)
		Copy Metadata:  		Copy Metadata ADS of selected files to clipboard
		Dump Metadata:  		Dump Metadata ADS of selected files to DOpus Output window
		Copy MediaInfo:  		Run MediaInfo anew and copy its output to clipboard
		Dump MediaInfo: 		Run MediaInfo anew and dump its output to DOpus Output window
		Estimate Bitrate:		Calculate bitrate using a list of target 'bitrate/pixel' values
		Toggle Essential:		Toggle a user-configurable list of 'essential' columns, can toggle DOpus columns
		Toggle Optional:		Toggle a user-configurable list of 'optional' columns, can toggle DOpus columns
		Toggle Other:   		Toggle a user-configurable list of 'other' columns, can toggle DOpus columns
		Toggle Verbose: 		Toggle a user-configurable list of 'verbose' columns, can toggle DOpus columns
		Validate Config:		Validate current user configuration

		A preconfigured menu button with all available actions is in distro (extract OSP if not visible)
	*/}.toString().substituteVars(),


    'Customization': function(){/*
		If you want to adjust the VCodec & ACodec output, see options:
			CODEC_APPEND_ADDINFO
			CODEC_USE_SHORT_VARIANT
			RESOLUTION_APPEND_VERTICAL
			LOOKUP_CODECS - as a LAST resort!

		LOOKUP_CODECS can be used to show identical info to FourCC if you want,
		but I tell you FourCC is pretty useless, it's not standardized nor filled/respected/recognized by many encoders/muxers.
		Suggested: turn on one or more of the 'HELPER' columns (Toggle Verbose action) to see what info you can use to identify a certain codec.

		If you want to rename the labels, you need to create a JSON file in the script directory;
		see Configuration->Reference section for more info.

		First check the formatted Metadata ADS, e.g. via Tooltip or Dump/Copy for a selected file.
		Some fields are already stored in ADS, but have no dedicated columns, and you'd need to create only a new column.
		Chances are very high that the information you need is already there (except FourCC, since it is a placebo BS field).

		If you want to add new buttons or columns based on available info, see OnInit() method
		and use one of the existing commands & fields as template; note few fields e.g. 'Available' does not use MultiCol!

		If the info you want is not in the ADS, copy/dump the MediaInfo output for the file.
		If MediaInfo does recognize the field you want, check OnME_Update() and add the field.
		HOWEVER,
		adding new fields means adding new fields to the structure stored in ADS
		and that means, you will need to update previously processed files with the ADS before you can use new fields!
		Please keep that in mind.

		Feature Requests:

		Requesting new fields/user option is ok, with some conditions:
			- No obscure fields, like encoder 'Format_Settings_Matrix', should be usable by many
			- My time allows it
			- You must supply a sample file and MediaInfo output of the sample, where the field(s) you want are marked clearly
	*/}.toString().substituteVars(),


    'Features': function(){/*
		Nearly constant scanning speed:
			File scanning speed does NOT depend on the individual file sizes
			since most containers store metadata at the front of the files and this is what MediaInfo parses.
			This means, the overall scanning duration depends more or less
			on the total number of files, not their sizes.
			The scanning speed depends on your disk's read speed of course
			but since SDD & NVMEs ubiquitous these days the speed is pretty acceptable.

		Dirty detection:
			If a file has been changed, e.g. overwritten by a program or its file-specific metadata (ID3, keywords, etc.)
			have been updated, this is automatically recognized and can be used to update its Metadata ADS.

		Progress window:
			"Pause/Resume" & "Abort" actions are supported fully,
			Skip seemed unnecessary to me since scanning speed per file is pretty high.

		Caching:

			This is an experimental feature and does not bring as much performance gain as you might expect from in-memory caching.

			Can be disabled via option.

			Disabling the cache and/or clearing it does not seem to make a huge impact; YMMV. Read on if you're curious.

			Any seen file, as long as the columns are visible, are cached in DOpus memory, ~1k of data per file.
			If the file path, name & last modification time have not been changed, the cached info is used.

			The Metadata info is principally processed using POJO/JSON,
			and my initial concern was reading JSON strings, converting them to POJO and back would be very time-consuming
			has been removed by another factor.

			DOpus seems to have a minimum overhead for script columns/fields ;
			i.e. even if you program the simplest fields/columns, updating the columns in a big directory
			still takes some time, although DOpus-internal fields are shown instantly.

			However, after much testing, it is safe to say,
			caching does not speed up the refresh speed of columns, since DOpus always has an overhead to display script fields/columns,
			even if they show a simple 'X' without calculating any information at all.

			The caching only reduces disk accesses, since NTFS ADS streams are practically separate files.

			The cached info is 'stringified JSONs', i.e. POJOs (Plain Old Javascript Objects) converted to readable format.
			This information is read and parsed (converted back to an in-memory POJO) whenever needed.
			Even if this takes some computing time, DOpus does not allow you to keep POJOs in memory for too long,
			these are removed from memory as soon as DOpus show 'Script Completed' in output window, even if you use Script.Vars object.
			Script.Vars can store simple types (integer, string, etc.) or DOpus objects, e.g. Vector & Map but not POJOs.
			Since it is too much hassle for me to convert all JS object accesses from object.field to DOPus Map() notation,
			I'm not planning to rewrite it and cache directly the parsed JSONs.
			Considering script-addins do not run faster than a certain speed, a rewrite would not bring much speed increase anyway.
			The only possible solution seems to be developing a DOpus plugin, which I'm not planning at the moment.

			Also added a parameter cache_max_items to limit the number of items in the cache based on timestamp of the item added to the cache
			but iterating over hash arrays based on an attribute other than the hash's key field turned out to be an unnecessary hassle
			too much pain for very little gain, so such a feature won't be implemented (again), just use the CLEAR CACHE command

		Multiple Video Tracks:
			You might wonder why "only 1st video stream" is mentioned elsewhere.
			The wonderful Matroska format allows to mux multiple video streams.
			But unlike files with multiple audio streams, no media player I know of can play such files in entirety,
			However, for the sake of completeness & OCD I have included them in the script.
	*/}.toString().substituteVars(),


    'How To Install': function(){/*
			- Download MediaInfo CLI version from: https://mediaarea.net/en/MediaInfo/Download/Windows
			- Put the .EXE file in any directory of your choice
			- Import the OSP file
			- Configure the script options in Preferences->Scripts section (click on underlined script name)
			- Add supplied commands in Customize->Commands->Script Commands to any toolbar/menu of your choice
			- Add new fields to your listers as columns or IntoTips/Tiles...

			Manual install:
			- Download & extract MediaInfo CLI version as above
			- Extract the OSP file (it's a ZIP)
			- Drag & drop / import the .JS file to your Script Addins
			- Copy the supplied icons folder under the script addins directory
			- Configure as above
			- Drag & drop the supplied .DCF buttons to a toolbar/menu of your choice
				or via Customize->Commands->Script Commands
	*/}.toString().substituteVars(),


    'DOpus Settings': function(){/*
		Although not required, this setting is HIGHLY RECOMMENDED:
			File Operations->Copy Attributes->Copy all NTFS streams: TRUE

		This ensures that if you copy a file the Metadata ADS is copied with it.
	*/}.toString().substituteVars(),


    'DISCLAIMER && LICENSE': function(){/*
		DISCLAIMER

		ALTHOUGH I HAVE EXTENSIVELY TESTED THIS WITH THOUSANDS OF FILES AND NTFS ADS STREAMS ARE BUILT-IN IN WINDOWS
		I DO NOT TAKE ANY RESPONSIBILITY WHATSOEVER THIS SCRIPT MIGHT DO TO YOUR FILES, COMPUTER, SANITY, CUTE BABY SEALS OR IMPENDING WORLD PEACE!
		IF YOU ARE UNSURE WHAT IT DOES, LEAVE NOW AND HAVE A GOOD DAY.
		IF YOU PROCEED, READ THE DESCRIPTIONS OF CONFIGURATION OPTIONS CAREFULLY
		AND TEST IT WITH COPIES OF YOUR FILES FIRST.
		EVEN IF EVERYTHING SEEMS TO WORK FOR MILLIONS OF FILES BUT ONE,
		I STILL DO NOT TAKE ANY RESPONSIBILITY.
		USE AT YOUR OWN RISK!


		LICENSE

		This script is subject to:
			${Global.SCRIPT_LICENSE}

		You are free to:
		* Share — copy and redistribute the material in any medium or format
		* Adapt — remix, transform, and build upon the material for any purpose, even commercially.

		Under the following terms:
		* Attribution — You must give appropriate credit, provide a link to the license, and indicate if changes were made.
		  You may do so in any reasonable manner, but not in any way that suggests the licensor endorses you or your use.
		* ShareAlike — If you remix, transform, or build upon the material, you must distribute your contributions under the same license as the original.

		No additional restrictions
		You may not apply legal terms or technological measures that legally restrict others from doing anything the license permits.

		Read full license at https://creativecommons.org/licenses/by-sa/4.0/


		CREDITS

		MediaInfo & MediaInfo icon: Copyright (c) 2002-2021 MediaArea.net SARL
		sprintf: https://hexmen.com/blog/2007/03/14/printf-sprintf/
		Android Material icons Copyright (c) Google https://material.io/resources/icons/
	*/}.toString().substituteVars(),


    'Version History': function(){/*
		v0.5:
			Initial Release

			Supported fields:
				Container Format, Video Format & Codec, Audio Format & Codec (only main audio track)
			Container/Total bitrate, Video Bitrate, Audio Bitrate (only main audio track)
				Channel count: 0 to 8 (shown as 0, 1.0, 2.0, 2.1, 4.0, 5.0, 5.1, 5.2, 7.1)
				Compression Mode: Lossy vs Lossless
				Bitrate Mode: VBR, ABR, CBR
				Column to show formatted JSON data stored in ADS
				Column to show raw JSON data stored in ADS
				'Multi-Audio' column to show if file has multiple audio tracks, e.g. for Matroska, etc.
				'Has Metadata' column which is never cached and separately calculated by checking if file has the special ADS attached to it
				'Dirty'  column to recognize if stored ADS data needs to be refreshed (compares file modification time against the one in ADS)
				Display Aspect Ratio (AR): some formats like Matroska and AVI(of all formats) support other ARs during playback, e.g. a 480x480 can be made to play as 480x360

			Available Commands:
				Updata Metadata (Insert/Update)
				Delete Metadata
				Copy ADS Metadata (to Clipboard)
				Dump MediaInfo (to DOpus Output)
				Copy MediaInfo (to Clipboard)

			Recognizes channel count & bitrate often more accurately than DOPus

			Supported filetypes:
				MKV, MP4, AVI, M4A, M4B, MP3, MP2, FLAC, AC3, AAC, DTS, TrueHD (Dolby Atmos)

			Customizable name for the ADS stream (see variable 'MetaStreamName')

			Option to keep the original file modification timestamp unchanged after adding/updating/deleting ADS data

			(Almost) constant parsing speed per file
				Most file formats store their metadata in their headers and MediaInfo uses this info.
				Further, this script uses only a subset of selected fields of MediaInfo output and uses ~1 KB per file.
				That is, even if your file is 1 MB or 50 GB, the processing speed is pretty constant and depends mostly on your disk speed.


		v0.51:
			Added progress windows; only Resume & Abort available (were mostly useless)


		v0.52:
			Auto-Refresh support added


		v0.6 (not released):

			Improved Progress Window:
				Initial implementation was very buggy and did not allow Resume or Abort, these do work now
				Skip does not make sense since update speed per file is constant and does not depend on file size at all


			Cache support:
				Seen, i.e. in any DOpus window listed, files can be cached in DOpus memory, ~1k per file.
				If the file path, name & last modification time have not been changed, the cached info is used.

				However, after much testing, it is safe to say,
				caching does not speed up the refresh speed of columns, since DOpus always has an overhead to display script fields/columns,
				even if they show a simple 'X' without calculating any information at all.

				The caching only reduces disk accesses, since NTFS ADS streams are practically separate files.

				The cached info is 'stringified JSONs' i.e. POJO Plain Old JavaScr.. erm JScript Objecta, converted to readable format.
				This information is read and parsed (converted back to an in-memory POJO) whenever needed.
				Even if this takes some computing time, DOpus does not allow you to keep POJOs in memory for too long,
				these are removed from memory as soon as DOpus show 'Script Completed' in output window, even if you use Script.Vars object.
				Script.Vars can store simple types (integer, string, etc.) or DOpus objects, e.g. Vector & Map but not POJOs.
				Since it is too much hassle for me to convert all JS object accesses from object.field to DOPus Map() notation,
				I'm not planning to rewrite it and cache directly the parsed JSONs.
				Considering script-addins do not run faster a certain speed, a rewrite would not bring much speed increase anyway.
				The only possible solution seems to be developing a DOpus plugin, which I'm not planning at the moment.

				Also added a parameter cache_max_items to limit the number of items in the cache based on timestamp of the item added to the cache
				but iterating over hash arrays based on an attribute other than the hash's key field turned out to be an unnecessary hassle
				too much pain for very little gain, so such a feature won't be implemented (again), just use the CLEAR CACHE command

			3 new commands:
				Clear Cache
				Dump ADS Metadata (to DOpus Output)
				Estimate Bitrates (based on various target bitrate per pixel values)

		Revised Codec List; if more information is available to distinguish very common codecs, this is shown, e.g. 'H264 (x264)' vs 'H264'
				Please keep in mind, the mythical FourCC is not reliable as you might think
				because it can be edited by quite a few editors and may not reflect what the used codec actually is.
				This happens often with the five million different versions of MPEG and various codecs.
				There was a time when many 3rd party codecs & tools used the codecIDs of Microsoft or Apple so that the common players can play such files.
				If you are unhappy with what is shown, see the variable 'CODECS_HASH' below and adjust as you wish.

			3 new helper columns to extend/adjust shown codecs list: ME Helper (Container), ME Helper (Video), ME Helper (Audio)

			3 new columns for number of Video, Audio & Text (Subtitles) streams

			1 new column for video dimension for codecs/files which DOpus misses (only M4S2 so far)
					defaults to MediaInfo value - sample code for MediaInfo value by default is available
					the values are 99.99% of the time identical, but this column does not show '320 x 240 x 24' but '320 x 240'

			1 new column for video raw Aspect Ratio (AR): Since Display AR was already available this seemed to be missing
					for most cases this is identical to Diplay AR but you can use compare this column to Display AR
					to find which videos have an explicitly set Display AR, or see next column

			1 new column for combined Aspect Ratio: Combines Display AR and Raw AR into one, appends (display_AR) to the column value if the 2 values do not match

			1 new column for video frame rate for codecs/files which DOpus misses (only FLV1 so far)
					defaults to DOpus value - sample code for MediaInfo value by default is available
					the values are mostly identical or very close (23.976 MediaInfo vs 23.98 DOpus)
					but expect some differences

			Probably more columns/fields, see the full list above

			Renamed 'ME Has Metadata' to 'ME Available'

			More file types & fields support:
				FLV, WEBM, Wave, Wave64, MP1, ALAC, TAK, TTA, DSD, AMR, Vorbis, AIFF, WavPack, WMA Lossy & Lossless
				some file types do support ReplayGain (at least Foobar2000 can tag & read them) but MediaInfo does not parse it, other fields are shown fine
			some uncommon file types are only partially supported only because of MediaInfo (raw AAC, raw DTS, MusePack)


		v0.7 (not released):

			Implemented/revised sorting
					empty values are shown properly before smallest alphabetical/numerical item

				Renamed columns & command names - BREAKS previous version

		v0.8 (not released):

			Custom Icons & OSP compatibility:
				Icons for:
					Add/Update, Clear Cache, Delete, Copy ADS Info, Dump ADS Info, Copy MediaInfo Output, Dump MediaInfo Output
					Estimate Bitrate, Settings, Column Add, Column Remove
					Gear (generic), Toggle On (generic), Toggle Off (generic) - not used at the moment


		v0.9:

			Even more columns

			Renamed Stream name - BREAKS previous versions

			DOpus Script Options

				All customizable options have been converted to user-editable options via Preferences -> Toolbars -> Scripts -> this script's name

			Completely revised grouping
				Something I didn't know you can auto-group a column by ALT-LEFTCLICK on the column header
				and ungroup by ALT-SHIFT-LEFTCLICK
				one can never truly explore all of DOpus!

			Added Info/About dialog
	*/}.toString().substituteVars(),


    'Future': function(){/*
		Github
			I'm planning to put everything to Github

		Extra Music fields?
			Since Foobar2000 & MP3Tag both are stellar programs, I see little need at the moment
			Probably a 'Cover' field, which is also used in MKVs and I plan to use it

		Images?
			DOpus already does a terrific job with images, I see little need at the moment
			Maybe a DPI field... but DOpus probably has it somewhere I haven't seen yet

		Community Feedback
			Although I have tested the script so far with many codecs I daily use or simply know of
			the world of multimedia is a jungle and some files, container/codec combinations are surely missed.

			Speed, how fast these columns are parsed/displayed, is another topic where I have no idea how it'd work for others.
			It's pretty fast on my machine, but I do have NVMEs & 12-core machine, so... no idea!

		v0.5:
			Initial release
		...
		many unreleased versions, see source if you're really that interested... I thought so.
		...
		v0.9:

			Even more columns

			Renamed Stream name - BREAKS previous versions

			DOpus Script Options

				All customizable options have been converted to user-editable options via Preferences -> Toolbars -> Scripts -> this script's name

			Completely revised grouping

			Add user configuration screen aka my personal Pandora's box

			Added Info/About dialog

			Customizable column headers
	*/}.toString().substituteVars(),
};

{
    delete HEREDOCS['How To Install']; // leave the one above as is, but does not make sense for About dialog, after user clearly installed the script
    delete HEREDOCS['Version History']; // leave the one above as is, but this one is too big and probably uninteresting to most users
    delete HEREDOCS['DOpus Settings']; // leave the one above as is, but this one is too big and probably uninteresting to most users
    delete HEREDOCS['Future']; // leave the one above as is, but this one is too big and probably uninteresting to most users
    var out = '';
    for (var hd in HEREDOCS) {
        HEREDOCS[hd] = hd +
			"\r\n" +
			HEREDOCS[hd]
			    .replace(/^(\s+?)\t(\S)/mg, '$1    $2')
				.replace(/^(\s+?)\t(\s*\S)/mg, '$1    $2')
			    .replace(/^(\s+?)\t(\s+\S)/mg, '$1    $2')
			    .replace(/^(\s+?)\t(\s+\S)/mg, '$1    $2')
			    .replace(/^(\s+?)\t(\s+\S)/mg, '$1    $2')
			//.replace(/^(\s+?)\t\b/mg, '$1    ')
			//.replace(/^\t\b/mg, '    ')
			// .replace(/^\t/mg, '    ')
			//.replace(/^    /mg, '')
			    .slice(14,-3)
			+ "\r\n";
        // DOpus.Output(hd);
        //out += HEREDOCS[hd];
    }
    // DOpus.Output(out);
    // DOpus.Output("\n" + HEREDOCS['WHAT IS THIS']);
    // DOpus.Output(JSON.stringify(HEREDOCS, null, 4));
}


// LOGGER object
var logger = (function() {
    var VALID_LEVELS = {
        NONE:    0,
        ERROR:   1,
        WARN:    2,
        NORMAL:  3,
        INFO:    4,
        VERBOSE: 5
    };
    var _level;
    function set_level (level) {
        DOpus.output('set_level() called with: ' + level);
        _level = typeof level === 'number' && level >= VALID_LEVELS.NONE && level <= VALID_LEVELS.VERBOSE ? level : _level; // if valid use new, if not use old
        if (typeof config !== 'undefined' && config.exists('DEBUG_LEVEL')) {
            config.set('DEBUG_LEVEL', _level);
        }
    }
    function get_level() {
        // very quick hack, only for DEBUG_LEVEL!
        if (typeof Script.config !== 'undefined' && typeof Script.config['DEBUG_LEVEL'] !== 'undefined') {
            _level = Script.config['DEBUG_LEVEL'];
        }
        if (typeof _level === 'undefined') {
            if (typeof config !== 'undefined' && config.exists('DEBUG_LEVEL')) {
                var cl = config.get('DEBUG_LEVEL');
                if (cl !== _level) { _level = cl; }
            } else {
                _level = VALID_LEVELS.ERROR;
            }
        }
        return _level;
    }
    function baseout (message, level) {
        get_level();
        if (level <= _level) DOpus.Output(message);
    }
    return {
        levels: VALID_LEVELS,

        force: function (message) {
            baseout(message, -1);
        },
        error: function (message) {
            baseout(message, this.levels.ERROR);
        },
        warn: function (message) {
            baseout(message, this.levels.WARN);
        },
        normal: function (message) {
            baseout(message, this.levels.NORMAL);
        },
        info: function (message) {
            baseout(message, this.levels.INFO);
        },
        verbose: function (message) {
            baseout(message, this.levels.VERBOSE);
        },
        setLevel: function (level) {
            DOpus.output('DOpus.setLevel() called with: ' + level);
            set_level(level);
        },
        getLevel: function () {
            return get_level();
        },
        getKeys: function () {
            var keys = [];
            for (var k in this.levels) {
                if (this.levels.hasOwnProperty(k)) {
                    keys.push(k);
                }
            }
            return keys;
        }
    }
}());


// CONFIG object
// based on Crockford's singleton
var config = (function () {
    /**
	 * private attribs & methods
	 */
    var pathValidator = DOpus.FSUtil;
    var SUPPORTED_TYPES = {
        BOOLEAN : 'BOOLEAN',
        STRING  : 'STRING',
        NUMBER  : 'NUMBER',
        PATH    : 'PATH',
        ARRAY   : 'ARRAY',
        POJO    : 'POJO',
        OBJECT  : 'OBJECT',
        REGEXP  : 'REGEXP',
        JSON    : 'JSON',
        FUNCTION: 'FUNCTION'
    };
    var ERROR_MODES = {
        NONE  : 'NONE',
        ERROR : 'ERROR',
        DIALOG: 'DIALOG'
    }
    var _count        = 0;
    var _vals         = {};
    var _types        = {};
    var _bindings     = {};
    var _error_mode   = ERROR_MODES.ERROR;

    // do not try to display any dialog during DOpus startup and until OnInit() finishes
    // i.e. no such shenanigans
    // config.setErrorMode(config.error.DIALOG);
    // config.addSimple('i', false, config.type.NUMBER, false);
    var dlg;
    /**
	 *
	 * @param {this.ERROR_MODES} em error mode
	 */
    function setErrorMode(em) {
        // do not call this with ERROR in main/global block otherwise script might fail to initialize
        if (!ERROR_MODES.hasOwnProperty(em)) {
            var msg = 'Error mode ' + em + ' is not supported';
            DOpus.Output(msg);
            showDialog(msg);
            return;
        }
        _error_mode = em;
    }
    function getErrorMode() {
        return _error_mode;
    }
    function showError(msg) {
        switch(_error_mode) {
            case ERROR_MODES.NONE:   return false;				// you can use this as: if(!addBoolean(...)) {/*error*/}
            case ERROR_MODES.ERROR:  throw new Error(msg);		// mainly for development
            case ERROR_MODES.DIALOG: return showDialog(msg);
        }
    }
    function showDialog(msg) {
        while(!DOpus.Listers && !DOpus.Listers(0) && !DOpus.Dlg()) {
            DOpus.Output('Waiting for DOpus to become visible');
            DOpus.Delay(500);
        }
        dlg		    = DOpus.Dlg();
        dlg.window  = DOpus.Listers(0);
        dlg.message = msg;
        dlg.title	= Global.SCRIPT_NAME + ' - Error' ;
        dlg.buttons	= 'OK';
        dlg.Show();
        dlg = null;
        return false;
    }

    function isValid(val, type) {
        var k;
        switch(type) {
            case SUPPORTED_TYPES.BOOLEAN:
                return typeof val === 'boolean';
            case SUPPORTED_TYPES.STRING:
                return typeof val === 'string';
            case SUPPORTED_TYPES.NUMBER:
                return typeof val === 'number';
            case SUPPORTED_TYPES.PATH:
                return typeof val === 'string' && pathValidator.Exists(pathValidator.Resolve(val));
            case SUPPORTED_TYPES.ARRAY:
                return typeof val === 'object' && val.length >= 0;
            case SUPPORTED_TYPES.POJO:
                // any object without functions
                if (typeof val !== 'object') {
                    return false;
                }
                for (k in val) {
                    if (typeof val[k] === 'function') { return false; }
                }
                return true;
            case SUPPORTED_TYPES.OBJECT:
                return typeof val === 'object';
            case SUPPORTED_TYPES.REGEXP:
                if (typeof val !== 'string' && typeof val !== 'object') {
                    return false;
                }
                try { eval('new RegExp(' + val + ');'); return true; } catch(e) { return false }
            case SUPPORTED_TYPES.JSON:
                try { JSON.parse(val); return true; } catch(e) { return false; }
            case SUPPORTED_TYPES.FUNCTION:
                return typeof val === 'function';
            default:
                return showError('isValid(): ' + type + ' is not supported');
        }
    }
    function addValueWithBinding(key, val, bindTo, type, bypassValidation) {
        if (_vals.hasOwnProperty(key)) {
            return showError(key + ' already exists');
        }
        if (!!!bypassValidation && !isValid(val, type)) {
            return showError('type ' + type + ' does not accept given value ' + val);
        }
        _count++;
        _vals[key]     = val;
        _types[key]    = type;
        _bindings[key] = bindTo;
        return true;
    }
    function getValue(key, autoGetDOpusValue) {
        // dammit JScript, not even default values!
        autoGetDOpusValue = typeof autoGetDOpusValue !== 'undefined' ? autoGetDOpusValue : true;

        var usingConfigVal = false;
        if (!_vals.hasOwnProperty(key)) {
            return showError(key + ' does not exist');
        }

        var valueToProbe;
        if (autoGetDOpusValue && typeof Script.config !== 'undefined' && typeof Script.config[_bindings[key]] !== 'undefined') {
            valueToProbe = Script.config[_bindings[key]];
            if (typeof valueToProbe === 'undefined' || valueToProbe === null) {
                return showError('Script config has no value for ' + key + ', check the binding: ' + _bindings[key]);
            }
            // TODO - DELETE
            // if (key === 'formats_regex_vbr' || key === 'formats_regex_lossless' || key === 'formats_regex_lossy') {
            // 	DOpus.Output('auto-converting ' + valueToProbe + ' to ' + _types[key]);
            // }
            valueToProbe = convert(valueToProbe, _types[key]);
            if (typeof valueToProbe === 'undefined') {
                return showError('Config value ' + _bindings[key] + ' is not valid');
            }
            usingConfigVal = true;
        } else {
            valueToProbe = _vals[key];
        }

        if (!isValid(valueToProbe, _types[key])) {
            return showError('Invalid value!\n\nKey:\t' + (usingConfigVal ? _bindings[key] : key) + '\nValue:\t' + valueToProbe + '\nUsing:\t' + (usingConfigVal ? 'User Config' : 'Default'));
        }
        return valueToProbe;
    }
    function SafeRegexpConvert(testString) {
        var res;
        if (typeof testString === 'string') {
            try { res = eval('new RegExp(' + testString + ');'); } catch(e) { /**/ }
        }
        return res;
    }
    function SafeJSONConvert(testString) {
        var res;
        if (typeof testString === 'string') {
            try { res = JSON.parse(testString); } catch(e) { /**/ }
        }
        return res;
    }
    function convert(val, type) {
        switch(type) {
            case SUPPORTED_TYPES.ARRAY:
            case SUPPORTED_TYPES.OBJECT:
            case SUPPORTED_TYPES.POJO:		return SafeJSONConvert(val);
            case SUPPORTED_TYPES.PATH:		return pathValidator.Resolve(val) + '';
            case SUPPORTED_TYPES.REGEXP:	return SafeRegexpConvert(val);
            default:						return val;
        }
    }
    function getType(key) {
        if (!_types.hasOwnProperty(key)) {
            return showError(key + ' does not exist');
        }
        return _types[key];
    }
    function getBinding(key) {
        if (!_bindings.hasOwnProperty(key)) {
            return showError(key + ' does not exist');
        }
        return _bindings[key];
    }
    function findBinding(bindTo) {
        for (var k in _bindings) {
            if (_bindings.hasOwnProperty(k) && _bindings[k] == bindTo) {
                return k;
            }
        }
        return false;
    }
    function setValue(key, val) {
        if (!_vals.hasOwnProperty(key)) {
            return showError(key + ' does not exist');
        }
        if (!isValid(val, _types[key])) {
            return showError(key + ' must have type ' + _types[key] + ', given: ' + typeof val);
        }
        // DOpus.Output('_types[key]: ' + _types[key]);
        // var _tmp;
        // if (_types[key] === SUPPORTED_TYPES.REGEXP && typeof val === 'string') {
        // 	DOpus.Output('regexp requested');
        // 	_tmp = SafeRegexpConvert(val);
        // 	if (_tmp === false) {
        // 		return showError(key + ' must have type ' + _types[key] + ', given value cannot be parsed as such');
        // 	}
        // 	val = _tmp;
        // } else if (_types[key] === SUPPORTED_TYPES.JSON && typeof val === 'string') {
        // 	DOpus.Output('json requested');
        // 	_tmp = SafeJSONConvert(val);
        // 	if (_tmp === false) {
        // 		return showError(key + ' must have type ' + _types[key] + ', given value cannot be parsed as such');
        // 	}
        // 	val = _tmp;
        // }
        _vals[key] = val;
    }
    function delValue(key) {
        if (!_vals.hasOwnProperty(key)) {
            return showError(key + ' does not exist');
        }
        var _tmp = _vals[key];
        count--;
        delete _vals[key];
        delete _types[key];
        delete _bindings[key];
        return _tmp;
    }
    function hasValue(key) {
        return _vals.hasOwnProperty(key);
    }

    /**
	 * public attribs & methods
	 */
    return {
        types: SUPPORTED_TYPES,
        modes: ERROR_MODES,
        /**
		 * @param {string} key config key
		 * @returns {boolean} true if key is valid
		 */
        exists: function(key) {
            return hasValue(key);
        },
        /**
		 * auto-checks the current Script.config value this key is bound to
		 * and returns the current value if valid and default value if invalid
		 *
		 * @param {string} key config key
		 * @param {boolean} autoGetDOpusValue get the Script.config value automatically, use false to get stored value
		 * @returns {any} config value
		 */
        get: function (key, autoGetDOpusValue) {
            return getValue(key, autoGetDOpusValue);
        },
        /**
		 * @param {string} key config key
		 * @param {any} val config value
		 */
        set: function (key, val) {
            setValue(key, val);
        },
        /**
		 *
		 * @param {string} key config key
		 * @param {boolean} val boolean
		 * @param {string} bindTo Script.config value to bind to
		 * @param {config.types} type value type
		 * @param {boolean} bypassValidation bypass validation
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addValueWithBinding: function(key, val, bindTo, type, bypassValidation) {
            return addValueWithBinding(key, val, bindTo, type, bypassValidation);
        },
        /**
		 * @param {string} key config key
		 * @param {boolean} val boolean
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addBoolean: function (key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.BOOLEAN);
        },
        /**
		 * @param {string} key config key
		 * @param {string} val string
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addString: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.STRING);
        },
        /**
		 * @param {string} key config key
		 * @param {number} val number (int, float...)
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addNumber: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.NUMBER);
        },
        /**
		 * given path is auto-resolved & checked for existence
		 *
		 * @param {string} key config key
		 * @param {string} val path
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addPath: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.PATH);
        },
        /**
		 * @param {string} key config key
		 * @param {array} val array
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addArray: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.ARRAY);
        },
        /**
		 * @param {string} key config key
		 * @param {object} val POJO, object without functions
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addPOJO: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.POJO);
        },
        /**
		 * @param {string} key config key
		 * @param {object} val object
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addObject: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.OBJECT);
        },
        /**
		 * @param {string} key config key
		 * @param {regexp} val regexp
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addRegexp: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.REGEXP);
        },
        /**
		 * @param {string} key config key
		 * @param {string} val JSON
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addJSON: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.JSON);
        },
        /**
		 * @param {string} key config key
		 * @param {function} val function
		 * @param {string} bindTo Script.config value to bind to
		 * @throws error (in ERROR_MODES.ERROR) if key already exists or value is invalid
		 */
        addFunction: function(key, val, bindTo) {
            return addValueWithBinding(key, val, bindTo, this.types.FUNCTION);
        },
        /**
		 * @param {string} key config key
		 * @throws error (in ERROR_MODES.ERROR) if key does not exist
		 */
        del: function (key) {
            delValue(key);
        },
        /**
		 * @returns {int} number of elements in the config
		 */
        getCount: function () {
            return _count;
        },
        /**
		 * @returns {array} keys in the config
		 */
        getKeys: function () {
            var keys = [];
            for (var k in _vals) if (_vals.hasOwnProperty(k)) keys.push(k);
            return keys;
        },
        /**
		 * @param {string} key config key
		 * @returns {config.types} type of value
		 */
        getType: function (key) {
            return getType(key);
        },
        /**
		 * @param {string} key config key
		 * @returns {string} bound Script.config key
		 */
        getBinding: function(key) {
            return getBinding(key);
        },
        /**
		 * @param {string} bindTo bound config variable name
		 * @returns {string|boolean} key name if found, false if not
		 */
        findBinding: function(bindTo) {
            return findBinding(bindTo);
        },
        /**
		 * @param {any} val config value - uses the type set by addXXX-methods
		 * @param {config.types} type - use config.types
		 * @returns {boolean} true if value is accepted
		 */
        validate: function (val, type) {
            return isValid(val, type, showError);
        },
        /**
		 * @returns {config.modes} current error mode
		 */
        getErrorMode: function() {
            return getErrorMode();
        },
        /**
		 * @param {string} em error mode, use config.modes
	 	 */
        setErrorMode: function(em) {
            setErrorMode(em);
        },
        /**
		 * @returns {string} stringified config
		 */
        toString: function() {
            return JSON.stringify(_vals, null, 4);
        },
        showError: showError,
        safeConvertToJSON: SafeJSONConvert,
        safeConvertToRegexp: SafeRegexpConvert,
    };
}());



// CONFIG - DEFAULT VALUES
{

    config.addNumber('DEBUG_LEVEL', logger.getLevel(), 'DEBUG_LEVEL');

    /**
	 * Name of the ADS stream, can be also used via "dir /:" or "type file:stream_name" commands
	 *
	 * ADJUST AS YOU SEE FIT
	 *
	 * WARNING:
	 * Make sure you use a long-term name for this stream.
	 * If you want to rename this after having processed some files,
	 * you should REMOVE all existing streams first (by calling the Remove command) before processing any file
	 * otherwise those streams will not be processed by this script and become orphans,
	 * and an army of thousands ghosts will haunt you for the rest of your life, you wouldn't like that mess
	 *
	 */
    config.addString('MetaStreamName', 'MExt_MediaInfo', 'META_STREAM_NAME');


    /**
	 * path to MediaInfo.exe, portable/CLI version can be downloaded from https://mediaarea.net/en/MediaInfo
	 *
	 * ADJUST AS YOU SEE FIT
	 *
	 * you need only the .exe file, no templates or alike are necessary as we use only built-in JSON output
	 */
    // config.addPath('mediainfo_path', '%gvdTool%\\MMedia\\MediaInfo\\MediaInfoXXX.exe', 'MEDIAINFO_PATH');
    // do not use addPath() but use addValueWithBinding() with bypass=true - otherwise people will get an error on initial installation
    config.addValueWithBinding('mediainfo_path', '%gvdTool%\\MMedia\\MediaInfo\\MediaInfo.exe', 'MEDIAINFO_PATH', config.types.PATH, true);
    // config.addString('ref_mediainfo_download_url', 'https://mediaarea.net/en/MediaInfo/Download/Windows', 'REF_MEDIAINFO_DOWNLOAD_URL');

    /**
	 * keep the original "last modified timestamp" after updating/deleting ADS; TRUE highly recommended
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    config.addBoolean('keep_orig_modts', true, 'KEEP_ORIG_MODTS');

    /**
	 * auto refresh lister after updating metadata
	 *
	 * ADJUST AS YOU SEE FIT
	 *
	 * please keep in mind, auto-refresh is not an automatic win
	 * if you select a single in big folder, the whole folder must be refreshed,
	 * i.e. all activated columns for all files need to be re-read
	 * which might be slow depending on your config and longer than you might have intended
	 *
	 * however, also keep in mind the read time per size DOES NOT DEPEND ON THE FILE SIZE
	 * the refresh time is relational to the NUMBER OF FILES and speed of your hdd/sdd/nvme/ramdisk/potato
	 */
    config.addBoolean('force_refresh_after_update', true, 'FORCE_REFRESH_AFTER_UPDATE');

    /**
	 * cache metadata JS objects in memory for unchanged files to speed up process (the gain remains questionable IMO)
	 *
	 * ADJUST AS YOU SEE FIT
	 *
	 * the information stored in ADS is usually small, typically ~1 KB per file
	 * the cache is used only and only if this script runs, i.e. automatically when any of the columns is visible (incl. in InfoTips/Tiles)
	 * or if you trigger a command
	 *
	 * to ensure the information is always up-to-date, the following is used:
	 *
	 * - caching disabled or file is not in cache => read from disk
	 * - file is already in cache => check file's last modification time against the timestamp in cached info
	 * 								   if different => read from disk
	 *   							   if same => use cache
	 * - UPDATE command executed (regardless of file was in cache or not) => add to/update cache
	 * - file is already in cache but DELETE command is executed => remove from cache
	 *
	 *
	 * CAVEATS:
	 * if you keep DOpus open for a long time and handle a lot of small files (only when any of these extra columns is shown),
	 * the memory usage of DOpus will increase
	 * this is not a 'memory leak' though, just the script doing what you tell it to: cache stuff :)
	 *
	 * to avoid high mem usage you can manually call the CLEARCACHE command via button, menu...
	 */
    config.addBoolean('cache_enabled', true, 'CACHE_ENABLED');


    // get list of all columns with ^\s+col(?!\.name).+$\n in a decent editor
    var fields_base_reference, fields_essential = [], fields_optional = [], fields_verbose = [], fields_other = [];
    // fields_base_reference = {
    fields_base_reference = function(){return{
        // This is just the help variable for your reference
        // (this is not valid JSON but makes it possible for me to show you more info)
        //
        // When you edit the real variable, just make sure it is valid JSON:
        // no single quotes, only double quotes, no comments and no trailing , in last element
        //
        // You can also use standard DOpus fields and mix it up with these fields
        // e.g. MExt_TotalDuration, mp3songlength, MExt_VResolution, picwidth, picheight...
        "MExt_HasMetadata"                : "essential",
        "MExt_NeedsUpdate"                : "essential",
        "MExt_TotalBitrate"               : "essential",
        "MExt_VideoCodec"                 : "essential",
        "MExt_VideoBitrate"               : "essential",
        "MExt_AudioCodec"                 : "essential",
        "MExt_AudioBitrate"               : "essential",
        "MExt_CombinedDuration"           : "essential",
        "MExt_VDimensions"                : "essential",
        "MExt_VResolution"                : "essential",
        "MExt_VFrameRate"                 : "essential",
        "MExt_VFrameCount"                : "essential",
        "MExt_VARCombined"                : "essential",
        "MExt_MultiAudio"                 : "essential",
        "MExt_AudioChannels"              : "essential",
        "MExt_AudioLang"                  : "essential",
        "MExt_AudioBitrateMode"           : "essential",
        "MExt_AudioCompressionMode"       : "essential",
        "MExt_HasReplayGain"              : "essential",
        "MExt_VBitratePerPixel"           : "essential",
        "MExt_SubtitleLang"               : "essential",

        "MExt_GrossByterate"              : "optional",
        "MExt_TotalDuration"              : "optional",
        "MExt_VideoDuration"              : "optional",
        "MExt_AudioDuration"              : "optional",
        "MExt_VARDisplay"                 : "optional",
        "MExt_VARRaw"                     : "optional",
        "MExt_VideoCount"                 : "optional",
        "MExt_AudioCount"                 : "optional",
        "MExt_TextCount"                  : "optional",
        "MExt_OthersCount"                : "optional",
        "MExt_VEncLibName"                : "optional",
        "MExt_VEncLib"                    : "optional",
        "MExt_VCodecID"                   : "optional",
        "MExt_ACodecID"                   : "optional",
        "MExt_AFormatVersion"             : "optional",
        "MExt_AProfile"                   : "optional",
        "MExt_EncoderApp"                 : "optional",
        "MExt_DateEncoded"                : "optional",
        "MExt_DateTagged"                 : "optional",

        "MExt_HelperContainer"            : "other",
        "MExt_HelperVideoCodec"           : "other",
        "MExt_HelperAudioCodec"           : "other",
        "MExt_CleanedUpName"              : "other",

        "MExt_ADSDataFormatted"           : "verbose",
        "MExt_ADSDataRaw"                 : "verbose"

        // do not put , in the last line
    }
    }.toString().slice(17, -3);
    config.addString('fields_base_reference', fields_base_reference.normalizeLeadingWhiteSpace(), 'REF_ALL_AVAILABLE_FIELDS');
    fields_base_reference = JSON.parse(fields_base_reference);
    for (var f in fields_base_reference) {
        switch(fields_base_reference[f]) {
            case 'essential':	fields_essential.push(f); break;
            case 'optional':	fields_optional.push(f); break;
            case 'other':		fields_other.push(f); break;
            case 'verbose':		fields_verbose.push(f); break;
        }
    }
    config.addArray('fields_essential', fields_essential, 'TOGGLEABLE_FIELDS_ESSENTIAL');
    config.addArray('fields_optional', fields_optional, 'TOGGLEABLE_FIELDS_OPTIONAL');
    config.addArray('fields_other', fields_other, 'TOGGLEABLE_FIELDS_OTHER');
    config.addArray('fields_verbose', fields_verbose, 'TOGGLEABLE_FIELDS_VERBOSE');

    config.addString('fields_essential_after', 'Comments', 'TOGGLEABLE_FIELDS_ESSENTIAL_AFTER');
    config.addString('fields_optional_after', 'MExt_SubtitleLang', 'TOGGLEABLE_FIELDS_OPTIONAL_AFTER');
    config.addString('fields_other_after', '', 'TOGGLEABLE_FIELDS_OTHER_AFTER');
    config.addString('fields_verbose_after', '', 'TOGGLEABLE_FIELDS_VERBOSE_AFTER');


    /**
	 * video resolution translation hash
	 *
	 * use SD, HD-Ready, HD, UHD, 4K, 8K, etc. if you like
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    var lookup_resolutions = function(){return{
        // This is just the help variable for your reference
        // (this is not valid JSON but makes it possible for me to show you more info)
        //
        // When you edit the real variable, just make sure it is valid JSON:
        // no single quotes, only double quotes, no comments and no trailing , in last element
        //
        // the comparison is always <= (less or equal), i.e. real duration <= config values below
        "240":      "240p",
        "360":      "360p",
        "480":      "480p",
        "576":      "576p",
        "720":      "720p",
        "1080":     "1080p",
        "2160":     "2160p",
        "4320":     "4320p"

        // do not put , in the last line
    }
    }.toString().slice(17, -3);
    JSON.stringify(JSON.parse(lookup_resolutions)); // test parseability on script load, do not remove
    config.addString('ref_lookup_resolutions', lookup_resolutions.normalizeLeadingWhiteSpace(), 'REF_LOOKUP_RESOLUTIONS');
    config.addPOJO('lookup_resolutions', JSON.parse(lookup_resolutions), 'LOOKUP_RESOLUTIONS');


    var lookup_duration_groups = function(){return{
        // This is just the help variable for your reference
        // (this is not valid JSON but makes it possible for me to show you more info)
        //
        // When you edit the real variable, just make sure it is valid JSON:
        // no single quotes, only double quotes, no comments and no trailing , in last element
        //
        // the comparison is always <= (less or equal), i.e. real duration <= config values below
        // if reported duration is 0 but the file definitely has audio e.g raw AAC/DTS/Atmos.., you can recognize this as well
        // \u2260 is the Unicode 'not equal' sign
        // if the file has no audio track at all, this is grouped automatically under 'No Audio', no need to define it here
        // Note that some values like 'Over 1h' depend on preceeding key's value
        "0":        " ≠00:00",
        "60":       "< 01:00",
        "120":      "01:00-02:00",
        "180":      "02:00-03:00",
        "240":      "03:00-04:00",
        "300":      "04:00-05:00",
        "600":      "05:00-10:00",
        "900":      "10:00-15:00",
        "1200":     "15:00-20:00",
        "1800":     "20:00-30:00",
        "3600":     "30:00-1:00:00",
        "5400":     "Over 1h",
        "7200":     "Over 1.5h",
        "10800":    "Over 2h",
        "999999":   "Over 3h"

        // do not put , in the last line
    }
    }.toString().slice(17, -3);
    JSON.stringify(JSON.parse(lookup_duration_groups)); // test parseability on script load, do not remove
    config.addString('ref_lookup_duration_groups', lookup_duration_groups.normalizeLeadingWhiteSpace(), 'REF_LOOKUP_DURATION_GROUPS');
    config.addPOJO('lookup_duration_groups', JSON.parse(lookup_duration_groups), 'LOOKUP_DURATION_GROUPS');

    /**
	 * video & audio codecs translation hash
	 *
	 * ADJUST AS YOU SEE FIT
	 *
	 * you might have to experiment a little bit to have the output suitable for your needs
	 * I've tried to keep them as close to DOpus as possible, but concerning how wild the MPEG specs are
	 * and how different encoders encode the videos, muxers set FourCC codes and other metainfo
	 * you might not always see what you see in another program, e.g. AviDemux might show it as DIVX and another program as MP42, etc.
	 */
    var lookup_codecs = function(){return{
        // This is just the help variable for your reference
        // (this is not valid JSON but makes it possible for me to show you more info)
        //
        // When you edit the real variable, just make sure it is valid JSON:
        // no single quotes, only double quotes, no comments and no trailing , in last element
        //
        // If there is an array with 2 elements on the right
        // you can switch to 2nd "short" version via script config.
        //
        // All fields below are first upper-cased then probed
        //
        // video probed in following order - activate the Helper columns if necessary
        // #1. CONTAINER_FORMAT - VIDEO_FORMAT - VIDEO_CODEC - ENC_LIBNAME
        // #2. CONTAINER_FORMAT - VIDEO_FORMAT - VIDEO_CODEC
        // #3. CONTAINER_FORMAT - VIDEO_FORMAT - ENC_LIBNAME
        // #4. CONTAINER_FORMAT - VIDEO_FORMAT
        // #5. CONTAINER_FORMAT - VIDEO_CODEC
        // #6. VIDEO_FORMAT - VIDEO_CODEC - ENC_LIBNAME
        // #7. VIDEO_FORMAT - VIDEO_CODEC
        // #8. VIDEO_FORMAT - ENC_LIBNAME
        // #9. VIDEO_FORMAT
        // #A. VIDEO_CODEC
        //
        // the more specific items, i.e. items near the top, should be preferred
        // but should be non-specific enough in case they are embedded in a movie file
        // that is why you should prefer AAC-MP4A-40-2 or AAC to MKA-AAC-MP4A-40-2 or MP4-AAC-MP4A-40-2
        // otherwise you would have to re-declare the same codecs over and over for every movie container
        // use one of the helper columns if you find something which is not shown to your liking
        "AV1-AV01"                                               : "AV1",
        "AV1"                                                    : "AV1",
        "AVC-AVC1-X264"                                          : ["H264 (x264)", "H264"],
        "AVC-H264-X264"                                          : ["H264 (x264)", "H264"],
        "AVC-V_MPEG4/ISO/AVC-X264"                               : ["H264 (x264)", "H264"],
        "AVC-X264"                                               : ["H264 (x264)", "H264"],
        "AVC-AVC1"                                               : "H264",
        "AVC"                                                    : "H264",
        "HEVC-HVC1-X265"                                         : ["H265 (x265)", "H265"],
        "HEVC-V_MPEGH/ISO/HEVC-X265"                             : ["H265 (x265)", "H265"],
        "HEVC-X265"                                              : ["H265 (x265)", "H265"],
        "HEVC-HVC1"                                              : "H265",
        "HEVC"                                                   : "H265",

        "HUFFYUV-HFYU"                                           : ["HuffYUV", "HFYU"],
        "HUFFYUV-V_MS/VFW/FOURCC / HFYU"                         : ["HuffYUV", "HFYU"],
        "HUFFYUV"                                                : ["HuffYUV", "HFYU"],

        "MPEG-4 VISUAL-DIV3"                                     : "Div3",
        "MPEG-4 VISUAL-V_MPEG4/MS/V3"                            : "Div3",
        "MPEG-4 VISUAL-DIVX"                                     : "DivX",
        "MPEG-4 VISUAL-V_MPEG4/ISO/ASP"                          : "DivX",
        "MPEG-4 VISUAL-DX50"                                     : "DivX5",
        "MPEG-4 VISUAL-MP4V-20"                                  : ["M4S2/MP4V", "MP4V"],
        "MPEG-4 VISUAL-MP42"                                     : "MP42",
        "MPEG-4 VISUAL-V_MS/VFW/FOURCC / MP42"                   : "MP42",
        "MPEG-4 VISUAL-V_MS/VFW/FOURCC / XVID"                   : "XviD",
        "MPEG-4 VISUAL-V_MS/VFW/FOURCC / DIV3"                   : "DivX",
        "MPEG-4 VISUAL-V_MS/VFW/FOURCC / DIVX"                   : "DivX",
        "MPEG-4 VISUAL-V_MS/VFW/FOURCC / DX50"                   : "DivX5",
        "MPEG-4 VISUAL-MP43"                                     : "MP43",
        "MPEG-4 VISUAL-XVID"                                     : "XviD",
        // "MPEG-4 VISUAL-DIV3"                                     : "DivX",
        "XVID-XVID"                                              : "XviD",
        "XVID"                                                   : "XviD",
        "MPEG VIDEO-MP4V-6A"                                     : "MPG1",
        "MPEG VIDEO-V_MPEG1"                                     : "MPG1",
        "MPEG VIDEO"                                             : "MPEG",
        "SORENSON SPARK"                                         : "FLV1",
        "VP6-VP6F"                                               : "VP6",
        "VP6"                                                    : "VP6",
        "VP8-VP8F"                                               : "VP8",
        "VP8"                                                    : "VP8",
        "VP9-V_VP9"                                              : "VP9",
        "VP9"                                                    : "VP9",
        "H.263-S263"                                             : "H263",
        "H.263"                                                  : "H263",
        "VC-1-WMV3"                                              : ["WMV3 (VC1)", "WMV3"],
        "VC-1"                                                   : "WMV3",
        "V_QUICKTIME"                                            : "",

        // All fields below are first upper-cased then probed
        //
        // audio probed in following order - activate the Helper columns if necessary
        // #1. CONTAINER_FORMAT - AUDIO_FORMAT - AUDIO_CODEC - AUDIO_FORMAT_VERSION - AUDIO_FORMAT_PROFILE - AUDIO_SETTINGS_MODE
        // #2. CONTAINER_FORMAT - AUDIO_FORMAT - AUDIO_CODEC - AUDIO_FORMAT_VERSION - AUDIO_FORMAT_PROFILE
        // #3. CONTAINER_FORMAT - AUDIO_FORMAT - AUDIO_CODEC - AUDIO_FORMAT_VERSION
        // #4. CONTAINER_FORMAT - AUDIO_FORMAT - AUDIO_CODEC
        // #5. CONTAINER_FORMAT - AUDIO_FORMAT
        // #6. AUDIO_FORMAT - AUDIO_CODEC - AUDIO_FORMAT_VERSION - AUDIO_FORMAT_PROFILE - AUDIO_SETTINGS_MODE
        // #7. AUDIO_FORMAT - AUDIO_CODEC - AUDIO_FORMAT_VERSION - AUDIO_FORMAT_PROFILE
        // #8. AUDIO_FORMAT - AUDIO_CODEC - AUDIO_FORMAT_VERSION
        // #9. AUDIO_FORMAT - AUDIO_CODEC
        // #A. AUDIO_FORMAT
        // #B. AUDIO_CODEC
        // #C. AUDIO_FORMAT - AUDIO_FORMAT_PROFILE  (mainly for MP3)

        "WMA-161-2"                                              : ["WMA (v9.2)", "WMA"],
        "WMA-162--PRO"                                           : ["WMA (v10 Pro)", "WMA"],
        "WMA-163--LOSSLESS"                                      : ["WMA (v9.2 Lossless)", "WMA"],
        "WMA-A"                                                  : ["WMA (v9 Voice)", "WMA"],
        "WMA"                                                    : ["WMA", "WMA"],

        "AC-3-A_EAC3"                                            : ["AC3 (Dolby Digital Plus)", "EAC3"],
        "E-AC-3-A_EAC3"                                          : ["AC3 (Dolby Digital Plus)", "EAC3"],

        "AC-3----DOLBY DIGITAL"                                  : ["AC3 (Dolby Digital)", "AC3"],
        "AC-3"                                                   : "AC3",

        "AIFF-PCM"                                               : ["PCM (AIFF)", "PCM"],

        "AMR-AMR---NARROW BAND"                                  : ["AMR (Narrow Band)", "AMR"],
        "AMR-SAMR---NARROW BAND"                                 : ["AMR (Narrow Band)", "AMR"],
        "AMR"                                                    : "AMR",

        "MONKEY'S AUDIO"                                         : ["Monkey's Audio", "APE"],
        "DSD"                                                    : "DSD",
        "FLAC"                                                   : "FLAC",
        "DTS"                                                    : "DTS",

        "ADTS-AAC"                                               : ["AAC (Raw)", "AAC"],
        "AAC-MP4A-40-2"                                          : "AAC",
        "AAC"                                                    : "AAC",

        "ALAC-ALAC"                                              : ["ALAC (Apple Lossless)","ALAC"],
        "ALAC"                                                   : ["ALAC (Apple Lossless)","ALAC"],

        "MLP FBA"                                                : ["Dolby TrueHD", "TrueHD"],
        "MUSEPACK SV8"                                           : ["Musepack", "MPC"],

        // MPEG audio is major PITA like MPEG video!
        "MPEG AUDIO--1-LAYER 2"                                  : ["MP2 (v1)", "MP2"],
        "MPEG AUDIO--2-LAYER 2"                                  : ["MP2 (v2)", "MP2"],
        "MPEG AUDIO--2.5-LAYER 2"                                : ["MP2 (v2.5)", "MP2"],
        "MPEG AUDIO-50-1-LAYER 2"                                : ["MP2 (v1)", "MP2"],
        "MPEG AUDIO-50-2-LAYER 2"                                : ["MP2 (v2)", "MP2"],
        "MPEG AUDIO-50-2.5-LAYER 2"                              : ["MP2 (v2.5)", "MP2"],
        "MPEG AUDIO-55-1-LAYER 2"                                : ["MP2 (v1)", "MP2"],
        "MPEG AUDIO-55-2-LAYER 2"                                : ["MP2 (v2)", "MP2"],
        "MPEG AUDIO-55-2.5-LAYER 2"                              : ["MP2 (v2.5)", "MP2"],
        "MPEG AUDIO-MP4A-69-1-LAYER 2"                           : ["MP2 (v1)", "MP2"],
        "MPEG AUDIO-MP4A-69-2-LAYER 2"                           : ["MP2 (v2)", "MP2"],
        "MPEG AUDIO-MP4A-69-2.5-LAYER 2"                         : ["MP2 (v2.5)", "MP2"],
        "MPEG AUDIO-MP4A-6B-1-LAYER 2"                           : ["MP2 (v1)", "MP2"],
        "MPEG AUDIO-MP4A-6B-2-LAYER 2"                           : ["MP2 (v2)", "MP2"],
        "MPEG AUDIO-MP4A-6B-2.5-LAYER 2"                         : ["MP2 (v2.5)", "MP2"],
        "MPEG AUDIO-A_MPEG/L3-1-LAYER 2"                         : ["MP2 (v1)", "MP2"],
        "MPEG AUDIO-A_MPEG/L3-2-LAYER 2"                         : ["MP2 (v2)", "MP2"],
        "MPEG AUDIO-A_MPEG/L3-2.5-LAYER 2"                       : ["MP2 (v2.5)", "MP2"],
        "MPEG AUDIO--1-LAYER 3"                                  : ["MP3 (v1)", "MP3"],
        "MPEG AUDIO--2-LAYER 3"                                  : ["MP3 (v2)", "MP3"],
        "MPEG AUDIO--2.5-LAYER 3"                                : ["MP3 (v2.5)", "MP3"],
        "MPEG AUDIO-50-1-LAYER 3"                                : ["MP3 (v1)", "MP3"],
        "MPEG AUDIO-50-2-LAYER 3"                                : ["MP3 (v2)", "MP3"],
        "MPEG AUDIO-50-2.5-LAYER 3"                              : ["MP3 (v2.5)", "MP3"],
        "MPEG AUDIO-55-1-LAYER 3"                                : ["MP3 (v1)", "MP3"],
        "MPEG AUDIO-55-2-LAYER 3"                                : ["MP3 (v2)", "MP3"],
        "MPEG AUDIO-55-2.5-LAYER 3"                              : ["MP3 (v2.5)", "MP3"],
        "MPEG AUDIO-MP4A-69-1-LAYER 3"                           : ["MP3 (v1)", "MP3"],
        "MPEG AUDIO-MP4A-69-2-LAYER 3"                           : ["MP3 (v2)", "MP3"],
        "MPEG AUDIO-MP4A-69-2.5-LAYER 3"                         : ["MP3 (v2.5)", "MP3"],
        "MPEG AUDIO-MP4A-6B-1-LAYER 3"                           : ["MP3 (v1)", "MP3"],
        "MPEG AUDIO-MP4A-6B-2-LAYER 3"                           : ["MP3 (v2)", "MP3"],
        "MPEG AUDIO-MP4A-6B-2.5-LAYER 3"                         : ["MP3 (v2.5)", "MP3"],
        "MPEG AUDIO-A_MPEG/L3-1-LAYER 3"                         : ["MP3 (v1)", "MP3"],
        "MPEG AUDIO-A_MPEG/L3-2-LAYER 3"                         : ["MP3 (v2)", "MP3"],
        "MPEG AUDIO-A_MPEG/L3-2.5-LAYER 3"                       : ["MP3 (v2.5)", "MP3"],

        // mega fallback
        // if you do not want all the MP1/MP2/MP3 details, comment the block above, and these 3 will be used
        "MPEG AUDIO-LAYER 1"                                     : "MP1",
        "MPEG AUDIO-LAYER 2"                                     : "MP2",
        "MPEG AUDIO-LAYER 3"                                     : "MP3",

        // Alternative, Multi-Channel aware definitions
        // WAVE64-PCM-00000003-0000-0010-8000-00AA00389B71--FLOAT'  : 'PCM (Wave64, 32bit, MultiCh)
        // WAVE64-PCM-3--FLOAT'                                     : 'PCM (Wave64, 32bit, 2.0)
        // WAVE-PCM-00000003-0000-0010-8000-00AA00389B71--FLOAT'    : 'PCM (Wave, 32bit, MultiCh)
        // WAVE-PCM-3--FLOAT'                                       : 'PCM (Wave, 32bit, 2.0)
        // WAVE64-PCM-1'                                            : 'PCM (Wave64, 16/24bit, 2.0)
        // WAVE-PCM-1'                                              : 'PCM (Wave, 16/24bit, 2.0)
        // WAVE64-PCM-00000001-0000-0010-8000-00AA00389B71'         : 'PCM (Wave64, 16/24bit, MultiCh)
        // WAVE-PCM-00000001-0000-0010-8000-00AA00389B71'           : 'PCM (Wave, 16/24bit, MultiCh)

        "WAVE64-PCM-00000003-0000-0010-8000-00AA00389B71--FLOAT" : ["PCM (Wave64, 32bit)", "PCM"],
        "WAVE64-PCM-3--FLOAT"                                    : ["PCM (Wave64, 32bit)", "PCM"],
        "WAVE-PCM-00000003-0000-0010-8000-00AA00389B71--FLOAT"   : ["PCM (Wave, 32bit)", "PCM"],
        "WAVE-PCM-3--FLOAT"                                      : ["PCM (Wave, 32bit)", "PCM"],
        "WAVE64-PCM-1"                                           : ["PCM (Wave64, 16/24bit)", "PCM"],
        "WAVE-PCM-1"                                             : ["PCM (Wave, 16/24bit)", "PCM"],
        "WAVE64-PCM-00000001-0000-0010-8000-00AA00389B71"        : ["PCM (Wave64, 16/24bit)", "PCM"],
        "WAVE-PCM-00000001-0000-0010-8000-00AA00389B71"          : ["PCM (Wave, 16/24bit)", "PCM"],
        "WAVE64-PCM"                                             : ["PCM (Wave64)", "PCM"],
        "WAVE-PCM"                                               : ["PCM (Wave)", "PCM"],
        "WAVE64"                                                 : ["PCM (Wave64)", "PCM"],
        "WAVE"                                                   : ["PCM (Wave)", "PCM"],

        "PCM"                                                    : "PCM",

        "OGG-OPUS"                                               : "Opus",
        "OPUS"                                                   : "Opus",
        "TAK-TAK"                                                : ["Tom's Audio", "TAK"],
        "TAK"                                                    : ["Tom's Audio", "TAK"],
        "TTA-TTA"                                                : ["TTA (TrueAudio)", "TTA"],
        "TTA"                                                    : ["TTA (TrueAudio)", "TTA"],
        "OGG-VORBIS"                                             : "Vorbis",
        "VORBIS"                                                 : "Vorbis",
        "WAVPACK-WAVPACK"                                        : ["WavPack", "WV"],
        "WAVPACK"                                                : ["WavPack", "WV"]

        // do not put , in the last line
    }
    }.toString().slice(17, -3);
    JSON.stringify(JSON.parse(lookup_codecs)); // test parseability on script load, do not remove
    config.addString('ref_lookup_codecs', lookup_codecs.normalizeLeadingWhiteSpace(), 'REF_LOOKUP_CODECS');
    config.addPOJO('lookup_codecs', JSON.parse(lookup_codecs), 'LOOKUP_CODECS');

    /**
	 * use short variants of codecs, found via LOOKUP_CODECS
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    config.addBoolean('codec_use_short_variant', false, 'CODEC_USE_SHORT_VARIANT');

    /**
	 * add container or codec-specific information to the container/video/audio codec fields automatically
	 * e.g. if an AAC file is encoded with 'LC SBR' it is shown as 'AAC (LC SPR)'
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    config.addBoolean('codec_append_addinfo', true, 'CODEC_APPEND_ADDINFO');

    /**
	 * append '(Vertical)' to video resolutions
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    config.addBoolean('resolution_append_vertical', true, 'RESOLUTION_APPEND_VERTICAL');

    /**
	 * Audio formats which do not store a VBR/CBR/ABR information separately but are VBR by definition
	 *
	 * do not touch
	 */
    config.addRegexp('formats_regex_vbr', new RegExp(/ALAC|Monkey's Audio|TAK|DSD/), 'FORMATS_REGEX_VBR');

    /**
	 * Audio formats which do not store a lossy/lossless information separately but are lossless by definition
	 *
	 * do not touch
	 */
    config.addRegexp('formats_regex_lossless', new RegExp(/ALAC|PCM|TTA|DSD/), 'FORMATS_REGEX_LOSSLESS');
    config.addRegexp('formats_regex_lossy', new RegExp(/AMR/), 'FORMATS_REGEX_LOSSY');

    /**
	 * audio channels translation hash
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    var lookup_channels = function(){return{
        // This is just the help variable for your reference
        // (this is not valid JSON but makes it possible for me to show you more info)
        //
        // When you edit the real variable, just make sure it is valid JSON:
        // no single quotes, only double quotes, no comments and no trailing , in last element
        //
        // alternative 1: '0': '0 (no audio)'
        // alternative 2: '0': '0 (n/a)'
        // alternative 3: '0': '0'
        // if you use   '0': ''   the value 0 will be shown as empty string as well, which impacts sorting
        "0": " ",

        "X": "> 0', // for formats like MusePack & raw DTS which report an audio track but not the channel count, sorted between 0 & 1",
        "X": "≠ 0', // for formats like MusePack & raw DTS which report an audio track but not the channel count, sorted between 0 & 1",
        "X": "≠0",

        "1": "1.0",
        "2": "2.0",
        "3": "2.1",
        "4": "4.0",
        "5": "5.0",
        "6": "5.1",
        "7": "5.2",
        "8": "7.1",
        "9": "7.2"

        // do not put , in the last line
    }
    }.toString().slice(17, -3);
    JSON.stringify(JSON.parse(lookup_channels)); // test parseability on script load, do not remove
    config.addString('ref_lookup_channels', lookup_channels.normalizeLeadingWhiteSpace(), 'REF_LOOKUP_CHANNELS');
    config.addPOJO('lookup_channels', JSON.parse(lookup_channels), 'LOOKUP_CHANNELS');

    /**
	 * directory in which temporary files (selected_files_name.JSON) are created
	 *
	 * ADJUST AS YOU SEE FIT
	 *
	 * NO trailing slashes & backslashes must be 'escaped', i.e. C:\MyTempDir\SubDir\ --> should be put as C:\\MyTempDir\\SubDir
	 *
	 * can include environment variables
	 */
    config.addPath('temp_files_dir', '%TEMP%', 'TEMP_FILES_DIR');

    /**
	 * EXPERIMENTAL
	 *
	 * get the MediaInfo output without using temporary files
	 */
    config.addBoolean('templess_mode', false, 'TEMPLESS_MODE');
    config.addNumber('templess_chunk_size', 32768, 'TEMPLESS_CHUNK_SIZE');

    /**
	 * external configuration file to adjust column headers
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    var config_file_dir_raw = '/dopusdata\\Script AddIns';
    var config_file_dir_resolved = DOpus.FSUtil.Resolve(config_file_dir_raw) + '\\';
    var config_file_name = Global.SCRIPT_NAME + '.json';
    var config_file_contents = function(){return{
        // To customize the column headers
        // create a file with the name: ${config_file_name}
        // under: ${config_file_dir_raw}
        // (usually: ${config_file_dir_resolved})
        //
        // DOpus does not allow column headers to be changed during runtime
        // but only during initialization.
        // That is, if you change these you have to restart DOpus or disable & enable the script.
        //
        // This is just the help variable for your reference
        // (this is not valid JSON but makes it possible for me to show you more info)
        //
        // When you edit the real variable, just make sure it is valid JSON:
        // no single quotes, only double quotes, no comments and no trailing , in last element
        //
        // prefix for all columns, e.g.
        // colPrefix: "",
        // colPrefix: ".",
        "colPrefix": "ME ",

        // use any replacement for any column
        // if a name for a column cannot be found, script default will be used
        // note all internal column names start with: ${Global.SCRIPT_PREFIX}
        //
        // replacement values must be valid JS string, i.e. quoted
        // e.g.
        "colRepl": {
            "MExt_CombinedDuration": "Duration",
            "MExt_GrossByterate": "Byterate",
            "MExt_MultiAudio": "Multiple Audio Tracks"
            // do not put , in the last line
        },

        // the following can be used to map vendor-specific fields
        // which MediaInfo usually puts under "Container.Extra" portion
        // to user-customizable fields. In order to find out how the fields
        // are called, you have to check out one of the helper fields and
        // look for the field "extra" with subitems under it.
        // this feature is supported only for Container.Extra, in order to
        // keep things simple, because there can be multiple audio/subtitle tracks,
        // or even multiple video tracks, but Extra blocks for those are rarely filled.
        "colExtra": {
            "com_apple_quicktime_model": "Camera Make",
            "com_apple_quicktime_software": "Camera Software"
            // do not put , in the last line
        }
        // do not put , in the last line
    }
    }.toString().slice(17, -3)
    .substituteVars();
    JSON.stringify(JSON.parse(config_file_contents)); // test parseability on script load, do not remove
    config.addString('ref_config_file', config_file_contents.normalizeLeadingWhiteSpace(), 'REF_CONFIG_FILE');



    /**
	 * name cleanup array
	 *
	 * use SD, HD-Ready, HD, UHD, 4K, 8K, etc. if you like
	 *
	 * ADJUST AS YOU SEE FIT
	 */
    var name_cleanup = function(){return{
        // This is just the help variable for your reference
        // (this is not valid JSON but makes it possible for me to show you more info)
        //
        // When you edit the real variable, just make sure it is valid JSON:
        // no single quotes, only double quotes, no comments and no trailing , in last element
        //
        // nameOnly applies to the name part of the file and extOnly to the extension
        "nameOnly": [
            // video codecs
            ["/x264|x265|H264|H265|DX50|DivX5|DivX4|DivX|XviD|VP9|VP6F|FLV1|HEVC|MPEG4|MP4V|MP42/ig",      ""],
            // audio codecs
            ["/AAC|AC3|DTS|MP3|OGG/ig",                                                          ""],
            // resolutions
            ["/240p|320p|360p|480p|540p|640p|720p|960p|1080p|1440p|2160p|4K|UHD/ig",             ""],
            ["/(_|\.)/g",                                                                        " " ],
            ["/(\d+).+$/g",                                                                      "$1" ],
            // normalize whitespace after all the replacements above
            ["/\s+/g",                                                                           " " ],
            ["/-\s*-/g",                                                                         "-" ]
        ],
        "extOnly": [
        ]
        // do not put , in the last line
    }
    }.toString().slice(17, -3);
    JSON.stringify(JSON.parse(name_cleanup)); // test parseability on script load, do not remove
    config.addString('ref_name_cleanup', name_cleanup.normalizeLeadingWhiteSpace(), 'REF_NAME_CLEANUP');
    config.addPOJO('name_cleanup', JSON.parse(name_cleanup), 'NAME_CLEANUP');



    // bindings for these 2 are unnecessary and ignored
    config.addString('config_file_dir_resolved', config_file_dir_resolved, '');
    config.addString('config_file_name', config_file_name, '');

    // do not touch this!
    config.addPOJO('ext_config_pojo', {});
}



// called by DOpus
// eslint-disable-next-line no-unused-vars
function OnInit(initData) {
    initData.name 			= Global.SCRIPT_NAME;
    initData.version 		= Global.SCRIPT_VERSION;
    initData.copyright 		= Global.SCRIPT_COPYRIGHT;
    initData.url 			= Global.SCRIPT_URL;
    initData.desc 			= Global.SCRIPT_DESC;
    initData.min_version 	= Global.SCRIPT_MIN_VERSION;
    initData.group			= Global.SCRIPT_GROUP;
    initData.log_prefix		= Global.SCRIPT_PREFIX;
    initData.default_enable = true;

    DOpus.ClearOutput();

    _readExternalConfig(initData);
    _initalizeConfigVars(initData);
    _initializeCommands(initData);
    _initializeColumns(initData);

    return false;
}



// internal method called by OnInit()
function addToConfigVar(initData, group, name, desc, value) {
    var cfg						= config.getBinding(name);
    initData.Config[cfg]		= value || config.get(name);
    // initData.config_desc(cfg)	= desc;
    // initData.config_groups(cfg)	= group;
    initData.config_desc.set(cfg, desc);
    initData.config_groups.set(cfg, group);
}
// internal method called by OnInit()
function addCommand(initData, name, method, template, icon, label, desc) {
    var cmd         = initData.AddCommand();
    cmd.name        = name;
    cmd.method      = method;
    cmd.template    = template;
    cmd.icon		= GetIcon(initData.file, icon);
    cmd.label		= label;
    cmd.desc        = desc;
}
// internal method called by OnInit()
function addColumn(initData, method, name, label, justify, autogroup, autorefresh, multicol) {
    var colPrefix   = 'ME ';
    var extConfig   = config.get('ext_config_pojo');
    var col         = initData.AddColumn();
    col.method      = method;
    col.name        = name;
    col.label       = (extConfig && extConfig.colPrefix || colPrefix) + (extConfig && typeof extConfig.colRepl === 'object' && extConfig.colRepl[name] || label);
    col.justify     = justify;
    col.autogroup   = autogroup;
    col.autorefresh = autorefresh;
    col.multicol    = multicol;
}

// internal method called by OnInit()
function _readExternalConfig(initData) {
    // var cfgpath = config.get('config_file_dir_resolved');
    // var cfgfile = config.get('config_file_name');
    var cfgpath = config_file_dir_resolved;
    var cfgfile = config_file_name;
    logger.force('Checking external config file under: ' + cfgpath + cfgfile);
    if (!DOpus.FSUtil.Exists( cfgpath + cfgfile)) {
        logger.force('...not found, skipping');
        return;
    }
    logger.force('...using external config file');
    var extcfgContents = ReadFile(cfgpath + cfgfile);
    // logger.force('extcfgContents:\n' + extcfgContents);

    // test parseability
    var cem = config.getErrorMode();
    config.setErrorMode(config.modes.DIALOG);
    try {
        config.set('ext_config_pojo', JSON.parse(extcfgContents));
        logger.force('...external config is valid JSON');
        // oExtConfigPOJO = JSON.parse(extcfgContents);
    } catch(e) {
        var err = 'External config file found but is not valid JSON, ignoring\n\nerror: ' + e.toString();
        config.showError(err);
        logger.force(err);
        config.setErrorMode(cem);
        return;
    }
}

// internal method called by OnInit()
function _initalizeConfigVars(initData) {

    initData.config_desc	= DOpus.Create.Map();
    initData.config_groups	= DOpus.Create.Map();
    var GROUP, _tmp;

    GROUP = 'zzDO NOT CHANGE UNLESS NECESSARY';
    addToConfigVar(initData, GROUP, 'MetaStreamName',
        'Name of NTFS stream (ADS) to use\nWARNING: DELETE existing ADS from all files before you change this, otherwise old streams will be orphaned'
    );


    GROUP = 'System';
    addToConfigVar(initData, GROUP, 'keep_orig_modts',
        'Keep the original "last modified timestamp" after updating/deleting ADS data\nMust be TRUE if you use \'Dirty/Needs Update\' column'
    );
    addToConfigVar(initData, GROUP, 'cache_enabled',
        'Enable in-memory cache\nRead About->Features for more info'
    );


    GROUP = 'Paths';
    addToConfigVar(initData, GROUP, 'mediainfo_path',
        'Path to MediaInfo.exe; folder aliases and %envvar% are auto-resolved\nportable/CLI version can be downloaded from https://mediaarea.net/en/MediaInfo/Download/Windows'
        , '/programfiles\\MediaInfo\\MediaInfo.exe'
    );
    addToConfigVar(initData, GROUP, 'temp_files_dir',
        'Temp dir for MediaInfo output files, folder aliases and %envvar% are auto-resolved\nTemp files are immediately deleted after a file has been processed\n'
    );
    addToConfigVar(initData, GROUP, 'templess_mode',
        'EXPERIMENTAL: Get the MediaInfo output without using temporary files.\nUse at your own risk, might not work!\n'
    );
    addToConfigVar(initData, GROUP, 'templess_chunk_size',
        'EXPERIMENTAL: Reading chunk size in Templess Mode; do not set too low!\nIf the value does not work, try 16384, 8192, 4096 and so on.\n'
    );



    GROUP = 'Listers';
    _tmp = DOpus.Create.Vector();
    _tmp.push_back(logger.getLevel());
    for (var i=0, keys=logger.getKeys(); i<keys.length; i++) {
        _tmp.push_back(keys[i]);
    }
    addToConfigVar(initData, GROUP, 'DEBUG_LEVEL',
        'How much information should be put to DOpus output window - Beware of anything above & incl. NORMAL, it might crash your DOpus!\nSome crucial messages or commands like Dump ADS, Dump MediaInfo, Estimate Bitrate are not affected',
        _tmp
    );
    addToConfigVar(initData, GROUP, 'force_refresh_after_update',
        'Force refresh lister after updating metadata (retains current selection)'
    );


    GROUP = 'Lookup/Translation - no ADS UPDATE necessary';
    addToConfigVar(initData, GROUP, 'lookup_resolutions',
        'Resolution lookup hash (JSON), use SD, HD-Ready, HD, UHD, 4K, 8K, etc. if you like\nMust be valid JSON, see reference section below',
        JSON.stringify(config.get('lookup_resolutions'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'lookup_duration_groups',
        'Durations lookup hash (JSON) - only for group columns, leave empty for normal value or use too short, too long, etc.; always sorted alphabetically\nMust be valid JSON, see reference section below',
        JSON.stringify(config.get('lookup_duration_groups'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'lookup_codecs',
        'Codecs lookup hash (JSON) - SEE SOURCE CODE for more info\nMust be valid JSON, see reference section below',
        JSON.stringify(config.get('lookup_codecs'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'lookup_channels',
        'Channel count lookup hash (JSON)\nMust be valid JSON, see reference section below',
        JSON.stringify(config.get('lookup_channels'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'codec_append_addinfo',
        'Append \'additional codec info\' after main codec info, e.g. AAC, DTS will be displayed as AAC (LC), DTS (XLL) etc.\nTo change the main codec info (e.g. MP3 (v1)--> MP3), try out CODEC_USE_SHORT_VARIANT=TRUE first, only then LOOKUP_CODECS'
    );
    addToConfigVar(initData, GROUP, 'codec_use_short_variant',
        'Use shorter version of codecs, e.g. APE instead of Monkey\'s Audio or MP3 instead of MP3 (v2)'
    );
    addToConfigVar(initData, GROUP, 'resolution_append_vertical',
        'Append (Vertical) to resolutions if height > width'
    );
    addToConfigVar(initData, GROUP, 'name_cleanup',
        'Name cleanup regexes for the \'Clean Name\' column (JSON)\nMust be valid JSON, see reference section below',
        JSON.stringify(config.get('name_cleanup'), null, 4).replace(/\n/mg, "\r\n")
    );


    GROUP = 'Lookup/Translation - ADS UPDATE necessary!';
    addToConfigVar(initData, GROUP, 'formats_regex_vbr',
        'Formats which do not report a VBR flag but are by definition VBR\nWARNING: Stored in ADS, i.e. if you change this, UPDATE is necessary, lister refresh will not suffice',
        config.get('formats_regex_vbr').toString()
    );
    addToConfigVar(initData, GROUP, 'formats_regex_lossless',
        'Formats which do not report a Lossless flag but are by definition Lossless\nWARNING: Stored in ADS, i.e. if you change this, UPDATE is necessary, lister refresh will not suffice',
        config.get('formats_regex_lossless').toString()
    );
    addToConfigVar(initData, GROUP, 'formats_regex_lossy',
        'Formats which do not report a Lossy flag but are by definition Lossy\nWARNING: Stored in ADS, i.e. if you change this, UPDATE is necessary, lister refresh will not suffice',
        config.get('formats_regex_lossy').toString()
    );


    GROUP = 'Toggleable Columns';
    addToConfigVar(initData, GROUP, 'fields_essential',
        'Essential Fields (sorted) - used exclusively by supplied action Toggle Essential, otherwise no impact on functionality\nYou can use DOpus fields as well, e.g. FOURCC, mp3songlength...',
        JSON.stringify(config.get('fields_essential'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'fields_essential_after',
        'Put essential fields after this column; if invalid/invisible then columns are added at the end'
    );
    addToConfigVar(initData, GROUP, 'fields_optional',
        'Optional fields (sorted) - used exclusively by supplied action Toggle Optional, otherwise no impact on functionality\nYou can use DOpus fields as well, e.g. FOURCC, mp3songlength...',
        JSON.stringify(config.get('fields_optional'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'fields_optional_after',
        'Put optional fields after this column; if invalid/invisible then columns are added at the end'
    );
    addToConfigVar(initData, GROUP, 'fields_other',
        'Other fields (sorted) - used exclusively by supplied action Toggle Optional, otherwise no impact on functionality\nYou can use DOpus fields as well, e.g. FOURCC, mp3songlength...',
        JSON.stringify(config.get('fields_other'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'fields_other_after',
        'Put other fields after this column; if invalid/invisible then columns are added at the end'
    );
    addToConfigVar(initData, GROUP, 'fields_verbose',
        'Verbose fields (sorted) - used exclusively by supplied action Toggle Optional, otherwise no impact on functionality\nYou can use DOpus fields as well, e.g. FOURCC, mp3songlength...',
        JSON.stringify(config.get('fields_verbose'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'fields_verbose_after',
        'Put verbose fields after this column; if invalid/invisible then columns are added at the end'
    );


    GROUP = 'zReference Only';	// first char is nbsp
    // addToConfigVar(initData, GROUP, 'ref_mediainfo_download_url',
    // 	'For your convenience ;), no affect at anything at all\nUse CLI version'
    // );
    var _prefix = 'For reference only - Changes have no effect at all\n';
    addToConfigVar(initData, GROUP, 'fields_base_reference',
        _prefix + 'List of all available columns; this list is used internally to calculate script defaults, change the TOGGLEABLE FIELDS options instead'
    );
    addToConfigVar(initData, GROUP, 'ref_lookup_resolutions',
        _prefix + 'Resolution lookup hash help'
    );
    addToConfigVar(initData, GROUP, 'ref_lookup_duration_groups',
        _prefix + 'Durations lookup hash help'
    );
    addToConfigVar(initData, GROUP, 'ref_lookup_codecs',
        _prefix + 'Codecs lookup hash help'
    );
    addToConfigVar(initData, GROUP, 'ref_lookup_channels',
        _prefix + 'Channel count lookup hash help'
    );
    addToConfigVar(initData, GROUP, 'ref_config_file',
        _prefix + 'External config file to customize column headers'
        // config.get('ref_config_file').replace(/\n/mg, "\r\n")
        //JSON.stringify(config.get('ref_config_file'), null, 4).replace(/\n/mg, "\r\n")
    );
    addToConfigVar(initData, GROUP, 'ref_name_cleanup',
        _prefix + 'Name cleanup regexes help'
    );

}
// internal method called by OnInit()
function _initializeCommands(initData) {
    /*
		Available icon names, used by GetIcon()
			Calculate
			ClearCache
			Add
			Remove
			Delete
			Info
			Settings
			Toggle_Off
			Toggle_On
			Update
	*/
    // function addCommand(initData, name, method, template, icon, label, desc)
    addCommand(initData,
        'ME_Update',
        'OnME_Update',
        'FILE/K',
        'AddUpdate',
        'MExt Update Metadata',
        'Update video and audio metadata (read by MediaInfo) in custom ADS stream');

    addCommand(initData,
        'ME_Delete',
        'OnME_Delete',
        'FILE/K',
        'Delete',
        'MExt Delete Metadata',
        'Delete video and audio metadata from custom ADS stream');

    addCommand(initData,
        'ME_ClearCache',
        'OnME_ClearCache',
        '',
        'ClearCache',
        'MExt Clear Cache',
        'Clear the in-memory cache');

    addCommand(initData,
        'ME_ADSDump',
        'OnME_ADSDump',
        'FILE/K',
        'Info',
        'MExt Dump Metadata',
        'Dump video and audio metadata stored in ADS to DOpus output window');

    addCommand(initData,
        'ME_ADSCopy',
        'OnME_ADSCopy',
        'FILE/K',
        'Info',
        'MExt Copy Metadata',
        'Copy video and audio metadata stored in ADS to clipboard');

    addCommand(initData,
        'ME_MediaInfoDump',
        'OnME_MediaInfoDump',
        'FILE/K',
        'Info',
        'MExt Dump MediaInfo',
        'Dump video and audio metadata (read by MediaInfo) to DOpus output window, without writing it to the ADS stream');

    addCommand(initData,
        'ME_MediaInfoCopy',
        'OnME_MediaInfoCopy',
        'FILE/K',
        'Info',
        'MExt Copy MediaInfo',
        'Copy video and audio metadata (read by MediaInfo) to Clipboard, without writing it to the ADS stream');

    addCommand(initData,
        'ME_EstimateBitrates',
        'OnME_EstimateBitrates',
        'FILE/K',
        'Calculate',
        'MExt Estimate Bitrate',
        'Calculate estimated bitrates for various target Bitrate/Pixel values');

    addCommand(initData,
        'ME_ToggleEssentialColumns',
        'OnME_ToggleEssentialColumns',
        '',
        'ToggleGroup1',
        'MExt Toggle Essential',
        'Toggle essential columns only');

    addCommand(initData,
        'ME_ToggleOptionalColumns',
        'OnME_ToggleOptionalColumns',
        '',
        'ToggleGroup2',
        'MExt Toggle Optional',
        'Toggle optional columns');

    addCommand(initData,
        'ME_ToggleOtherColumns',
        'OnME_ToggleOtherColumns',
        '',
        'ToggleGroup3',
        'MExt Toggle Other',
        'Toggle other columns');

    addCommand(initData,
        'ME_ToggleVerboseColumns',
        'OnME_ToggleVerboseColumns',
        '',
        'ToggleGroup4',
        'MExt Toggle Verbose',
        'Toggle verbose columns');

    addCommand(initData,
        'ME_ConfigValidate',
        'OnME_ConfigValidate',
        'DEBUG/S,SHOWVALUES/S',
        'Settings',
        'MExt Validate Config',
        'Validate current configuration');

    addCommand(initData,
        'ME_TestMethod1',
        'OnME_TestMethod1',
        '',
        'Settings',
        'MExt Test Method 1',
        'Test Method 1');

    addCommand(initData,
        'ME_TestMethod2',
        'OnME_TestMethod2',
        '',
        'Settings',
        'MExt Test Method 2',
        'Test Method 2');
}
// internal method called by OnInit()
function _initializeColumns(initData) {
    // function addColumn(initData, method, name, label, justify, autogroup, autorefresh, multicol)

    // this column is kept separate, no multicol
    addColumn(initData,
        'OnMExt_HasMetadata',
        'MExt_HasMetadata',
        'Available',
        'right', false, true, false);

    // all multicol below
    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_NeedsUpdate',
        'Dirty',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VideoCount',
        'VCount',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioCount',
        'ACount',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_TextCount',
        'TCount',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_OthersCount',
        'OCount',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VideoCodec',
        'VCodec',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VideoBitrate',
        'VBitrate',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_TotalBitrate',
        'TBitrate',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioCodec',
        'ACodec',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioBitrate',
        'ABitrate',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_TotalDuration',
        'TDuration',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VideoDuration',
        'VDuration',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioDuration',
        'ADuration',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_CombinedDuration',
        'Duration (Combined)',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_MultiAudio',
        'Multi-Audio',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioChannels',
        'Audio Channels',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioLang',
        'Audio Language',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VARDisplay',
        'Aspect Ratio (Display)',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VARRaw',
        'Aspect Ratio (Raw)',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VARCombined',
        'Aspect Ratio (Combined)',
        'right', false, true, true);

    // justified to right, it is basically a number-based field, left as DOpus uses, does not make sense to me
    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VDimensions',
        'Dimensions',
        'right', false, true, true);

    // justified to right, it is basically a number-based field, left as DOpus uses, does not make sense to me
    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VResolution',
        'Resolution',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VFrameRate',
        'Frame Rate',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VFrameCount',
        'Frame Count',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioBitrateMode',
        'ABitrate Mode',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AudioCompressionMode',
        'Audio Compression Mode',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_HasReplayGain',
        'RG',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_SubtitleLang',
        'Subtitle Lang',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_GrossByterate',
        'Gross KBps',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VBitratePerPixel',
        'VBitrate/Pixel',
        'right', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VEncLib',
        'Encoded Library',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VEncLibName',
        'Encoded Library Name',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_VCodecID',
        'Video Codec ID',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_ACodecID',
        'Audio Codec ID',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AFormatVersion',
        'Audio Format Version',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_AProfile',
        'Audio Format Profile',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_EncoderApp',
        'Container Encoder App',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_DateEncoded',
        'Container Encoded Date',
        'left', false, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_DateTagged',
        'Container Tagged Date',
        'left', false, true, true);


    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_HelperContainer',
        'Helper (Container)',
        'left', true, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_HelperVideoCodec',
        'Helper (VCodec)',
        'left', true, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_HelperAudioCodec',
        'Helper (ACodec)',
        'left', true, true, true);


    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_CleanedUpName',
        'Helper (CleanName)',
        'left', true, true, true);


    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_ADSDataFormatted',
        'ADSData (Formatted)',
        'left', true, true, true);

    addColumn(initData,
        'OnMExt_MultiColRead',
        'MExt_ADSDataRaw',
        'ADSData (Raw)',
        'left', true, true, true);


    // add user-customizable columns from Container.Extra fields
    // e.g.
    // "colExtra": {
    //   "com_apple_quicktime_model": "Camera Make"
    // }
    var extConfig = config.get('ext_config_pojo');
    if (typeof extConfig !== 'undefined' && typeof extConfig.colExtra === 'object') {
        /** @type {DOpusMap} */
        var colExtraMap = DOpus.create().map();
        for(var ecitem in extConfig.colExtra) {
            DOpus.output('adding custom column: ' + ecitem);
            addColumn(initData,
                'OnMExt_MultiColRead',
<<<<<<< HEAD
                'MExt_' + ecitem,
=======
                ecitem,
>>>>>>> b8370f1fc2d47fc98439d223376699b1629d4f27
                extConfig.colExtra[ecitem],
                'left', true, true, true);
            colExtraMap.set(ecitem, extConfig.colExtra[ecitem]);
        }
        initData.vars.set('colExtraMap', colExtraMap);
    }
}



// called by DOpus
// eslint-disable-next-line no-unused-vars
function OnScriptConfigChange(cfgChangeData) {
    logger.normal('OnScriptConfigChange() -- started');

    var oldErrorMode = config.getErrorMode();
    config.setErrorMode(config.modes.DIALOG);
    for (var i = 0; i < cfgChangeData.changed.count; i++) {
        var _tmp, err = '';
        var cfgkey = cfgChangeData.changed[i];
        var cfgval = Script.config[cfgkey];
        logger.info('processing config item: ' + cfgkey);

        DOpus.Output('Found key from binding: ' + config.findBinding(cfgkey));
        switch(cfgkey) {

            // case 'TOGGLEABLE_FIELDS_ESSENTIAL':
            // case 'TOGGLEABLE_FIELDS_OPTIONAL':
            // case 'TOGGLEABLE_FIELDS_OTHER':
            // case 'TOGGLEABLE_FIELDS_VERBOSE':
            // 	logger.verbose('config item (JSON): ' + cfgkey);
            // 	logger.error('config item (JSON): ' + cfgkey);
            // 	if (!_tmp) { err = 'Cannot parse invalid JSON: ' + cfgkey; break; }
            // 	switch(cfgkey) {
            // 		// case 'TOGGLEABLE_FIELDS_ESSENTIAL':	fields_essential = _tmp; break;
            // 		// case 'TOGGLEABLE_FIELDS_OPTIONAL':	fields_optional = _tmp; break;
            // 		// case 'TOGGLEABLE_FIELDS_OTHER':		fields_other = _tmp; break;
            // 		// case 'TOGGLEABLE_FIELDS_VERBOSE':	fields_verbose = _tmp; break;
            // 		case 'TOGGLEABLE_FIELDS_ESSENTIAL':	config.set('fields_essential', JSON.parse(cfgval)); break;
            // 		case 'TOGGLEABLE_FIELDS_OPTIONAL':	config.set('fields_optional', JSON.parse(cfgval)); break;
            // 		case 'TOGGLEABLE_FIELDS_OTHER':		config.set('fields_other', JSON.parse(cfgval)); break;
            // 		case 'TOGGLEABLE_FIELDS_VERBOSE':	config.set('fields_verbose', JSON.parse(cfgval)); break;
            // 	}
            // case 'FORMATS_REGEX_VBR':
            // case 'FORMATS_REGEX_LOSSLESS':
            // case 'FORMATS_REGEX_LOSSY':
            // 	_tmp = DELETE_____SafeRegexpCheck(cfgval);
            // 	if (!_tmp) { err = 'Cannot parse invalid RegExp: ' + cfgkey; break; }
            // 	switch(cfgkey) {
            // 		case 'FORMATS_REGEX_VBR':			formats_regex_vbr = _tmp; break;
            // 		case 'FORMATS_REGEX_LOSSLESS':		formats_regex_lossless = _tmp; break;
            // 		case 'FORMATS_REGEX_LOSSY':			formats_regex_lossy = _tmp; break;
            // 	}

            // case 'codec_append_addinfo':	codec_append_addinfo = cfgval; break;
            // case 'KEEP_ORIG_MODTS':						keep_orig_modts = cfgval; break;
            // case 'CACHE_ENABLED':						cache_enabled = cfgval; break;
            // case 'DEBUG_LEVEL':							DEBUG_LEVEL = cfgval; break;
            // case 'FORCE_REFRESH_AFTER_UPDATE':			force_refresh_after_update = cfgval; break;

            case 'DEBUG_LEVEL':
                logger.setLevel(cfgval);
                // config.set('DEBUG_LEVEL', cfgval); // I don't like it
                break;

            case 'META_STREAM_NAME':
                if(!cfgval || !config.validate(cfgval, config.types.STRING)) { err = 'Stream name cannot be empty or invalid'; break; }
                var dlgConfirm		= DOpus.dlg();
                dlgConfirm.message  = '\n\nIMPORTANT\n\nPlease make sure that you first remove existing ADS data\nfrom all files you processed so far\nusing the old name and leave no orphan streams.\n\nAn UPDATE of files will not remove the old streams,\nsince the new name will be in effect.\n\nIf you get stuck with orphans, you can use\n**Sysinternals Streams** tool,\nbut it\'ll also delete other streams you might want to keep.\n\n';
                dlgConfirm.title	= Global.SCRIPT_NAME + ' - Confirm' ;
                dlgConfirm.buttons	= 'OK';
                dlgConfirm.show();
                config.set('MetaStreamName', cfgval);
                break;


                // case 'TOGGLEABLE_FIELDS_ESSENTIAL_AFTER':	config.set('fields_essential_after', cfgval); break;
                // case 'TOGGLEABLE_FIELDS_OPTIONAL_AFTER':	config.set('fields_optional_after', cfgval); break;
                // case 'TOGGLEABLE_FIELDS_OTHER_AFTER':		config.set('fields_other_after', cfgval); break;
                // case 'TOGGLEABLE_FIELDS_VERBOSE_AFTER':	config.set('fields_verbose_after', cfgval); break;
                // case 'FORMATS_REGEX_VBR':			config.set('formats_regex_vbr', cfgval); break;
                // case 'FORMATS_REGEX_LOSSLESS':		config.set('formats_regex_lossless', cfgval); break;
                // case 'FORMATS_REGEX_LOSSY':			config.set('formats_regex_lossy', cfgval); break;

            case 'FORMATS_REGEX_VBR':
            case 'FORMATS_REGEX_LOSSLESS':
            case 'FORMATS_REGEX_LOSSY':
                _tmp = config.safeConvertToRegexp(cfgval);
                if (_tmp === false || typeof _tmp === 'undefined') {
                    err = 'Invalid Regexp in ' + cfgkey; break;
                }
                DOpus.Output('_tmp: ' + _tmp + ', type: ' + typeof _tmp);
                switch(cfgkey) {
                    case 'FORMATS_REGEX_VBR':			config.set('formats_regex_vbr', _tmp); break;
                    case 'FORMATS_REGEX_LOSSLESS':		config.set('formats_regex_lossless', _tmp); break;
                    case 'FORMATS_REGEX_LOSSY':			config.set('formats_regex_lossy', _tmp); break;
                }
                break;

            case 'TOGGLEABLE_FIELDS_ESSENTIAL':
                config.safeConvertToJSON(cfgval) && config.set('fields_essential', JSON.parse(cfgval)); break;
            case 'TOGGLEABLE_FIELDS_OPTIONAL':
                config.safeConvertToJSON(cfgval) && config.set('fields_optional', JSON.parse(cfgval)); break;
            case 'TOGGLEABLE_FIELDS_OTHER':
                config.safeConvertToJSON(cfgval) && config.set('fields_other', JSON.parse(cfgval)); break;
            case 'TOGGLEABLE_FIELDS_VERBOSE':
                config.safeConvertToJSON(cfgval) && config.set('fields_verbose', JSON.parse(cfgval)); break;

            case 'NAME_CLEANUP':
                config.safeConvertToJSON(cfgval) && config.set('name_cleanup', JSON.parse(cfgval)); break;


                // case 'MEDIAINFO_PATH':
                // 	if (!cfgval) { err = 'Path to MediaInfo must be configured'; break; }
                // 	_tmp = DOpus.FSUtil.Resolve(cfgval);
                // 	if(!DOpus.FSUtil.Exists(_tmp)) { err = 'Path to MediaInfo is invalid: ' + _tmp; break; }
                // 	config.set('mediainfo_path', cfgval);
                // 	break;

                // case 'TEMP_FILES_DIR':
                // 	if (!cfgval) { err = 'Temp directory must be configured'; break; }
                // 	// _tmp = DOpus.FSUtil.Resolve(cfgval);
                // 	// if(!DOpus.FSUtil.Exists(_tmp)) { err = 'Temp directory is invalid: ' + _tmp; break; }
                // 	// temp_files_dir = cfgval;
                // 	if(!config.validate('temp_files_dir', cfgval)) { err = 'Temp directory is invalid: ' + _tmp; break; }
                // 	config.set('temp_files_dir', cfgval);
                // 	break;

            // case 'REF_MEDIAINFO_DOWNLOAD_URL':
            // default:
            // 	// nothing
            default:
                var boundConfigKey = config.findBinding(cfgkey);
                if (boundConfigKey !== false) {
                    logger.force('Updating config value -- key: ' + boundConfigKey + ', bindTo: ' + cfgkey + ', new val: ' + cfgval);
                    config.set(boundConfigKey, cfgval);
                }
        }
        if (err) {
            logger.error(err);
            var dlg		= DOpus.Dlg();
            dlg.message = err;
            dlg.title	= Global.SCRIPT_NAME + ' - Error' ;
            dlg.buttons	= 'OK';
            ret = dlg.show;
            continue;
        }
    }
    validateConfig();
    config.setErrorMode(oldErrorMode);
    logger.normal('OnScriptConfigChange() -- finished');
}

// called by DOpus
// eslint-disable-next-line no-unused-vars
function OnAboutScript(aboutData) {
    var buttons = [];
    for(var hd in HEREDOCS) {
        buttons.push(hd);
    }
    buttons.push('Exit');

    var dlg		= DOpus.Dlg;
    dlg.window	= aboutData.window;
    dlg.title	= Global.SCRIPT_NAME + ' ' + Global.SCRIPT_VERSION + ' ' + Global.SCRIPT_DATE + ' ' + Global.SCRIPT_COPYRIGHT;
    dlg.buttons	= buttons.join('|');
    // dlg.icon	= GetIcon(Script.file, 'Info'); // bummer does not work
    dlg.icon	= 'info';

    logger.verbose('OnAboutScript -- Buttons: ' + buttons.join('|'));

    var buttonID = 1;
    while (buttonID) {
        if(buttonID == 1 )	dlg.message = HEREDOCS[buttons[0]];		dlg.title[buttons[0]];
        if(buttonID == 2 )	dlg.message = HEREDOCS[buttons[1]];		dlg.title[buttons[1]];
        if(buttonID == 3 )	dlg.message = HEREDOCS[buttons[2]];		dlg.title[buttons[2]];
        if(buttonID == 4 )	dlg.message = HEREDOCS[buttons[3]];		dlg.title[buttons[3]];
        if(buttonID == 5 )	dlg.message = HEREDOCS[buttons[4]];		dlg.title[buttons[4]];
        if(buttonID == 6 )	dlg.message = HEREDOCS[buttons[5]];		dlg.title[buttons[5]];
        if(buttonID == 7 )	dlg.message = HEREDOCS[buttons[6]];		dlg.title[buttons[6]];
        if(buttonID == 8 )	dlg.message = HEREDOCS[buttons[7]];		dlg.title[buttons[7]];
        if(buttonID == 9 )	dlg.message = HEREDOCS[buttons[8]];		dlg.title[buttons[8]];
        if(buttonID == 10)	dlg.message = HEREDOCS[buttons[9]];		dlg.title[buttons[9]];
        if(buttonID == 11)	dlg.message = HEREDOCS[buttons[10]];	dlg.title[buttons[10]];
        if(buttonID == 12)	dlg.message = HEREDOCS[buttons[11]];	dlg.title[buttons[11]];
        if(buttonID == 13)	dlg.message = HEREDOCS[buttons[12]];	dlg.title[buttons[12]];
        if(buttonID == 14)	dlg.message = HEREDOCS[buttons[13]];	dlg.title[buttons[13]];
        buttonID = dlg.Show();
        logger.verbose('OnAboutScript -- Clicked button code: ' + buttonID);
    }
    return false;
}


// called by 'Has Metadata' column
// eslint-disable-next-line no-unused-vars
function OnMExt_HasMetadata(scriptColData) {
    var selected_item = scriptColData.item;
    if (selected_item.is_dir || selected_item.is_reparse || selected_item.is_junction || selected_item.is_symlink) {
        return;
    }
    logger.normal("...Processing " + selected_item.name);
    // var fh = DOpus.FSUtil.OpenFile(selected_item.realpath + ':' + config.get('MetaStreamName')); // default read mode
    // var res = fh.error !== 0 ? false : true;
    // fh.Close();
    var res = DOpus.FSUtil.Exists(selected_item.realpath + ':' + config.get('MetaStreamName'));
    scriptColData.value = res ? 'Yes' : 'No';
    scriptColData.group = 'Has Metadata: ' + scriptColData.value;
    return res;
}

// called by all other columns than 'Has Metadata'
// eslint-disable-next-line no-unused-vars
function OnMExt_MultiColRead(scriptColData) {
    var ts1 = new Date().getTime();

    var selected_item   = scriptColData.item;
    if (selected_item.is_dir || selected_item.is_reparse || selected_item.is_junction || selected_item.is_symlink ) {
        return;
    }
    logger.normal('...Processing ' + selected_item.name);

    // get tags object
    var item_props = ReadMetadataADS(selected_item);
    if (item_props === false || typeof item_props === 'undefined' || !isObject(item_props)) {
        logger.normal(selected_item.name + ': Metadata does not exist or INVALID');
        return;
    }

    var _vcodec, _acodec, _resolution;
    var colExtraMap = Script.vars.get('colExtraMap');

    // iterate over requested columns
    for (var e = new Enumerator(scriptColData.columns); !e.atEnd(); e.moveNext()) {
        var key = e.item();

        var outstr = '', _tmp0, _tmp1, _tmp2, _tmp3, _tmp4, _tmp5, _tmp6, _tmp7;
        switch(key) {
            case 'MExt_NeedsUpdate':
                if (!config.get('keep_orig_modts')) {
                    scriptColData.columns(key).value = 'meaningless if KEEP_ORIG_MODTS is set to false';
                    break;
                }
                outstr = new Date(selected_item.modify).valueOf() === item_props.last_modify ? 0 : 1;
                logger.verbose('Old: ' + new Date(selected_item.modify).valueOf() + ' <> ' + item_props.last_modify);
                scriptColData.columns(key).group = 'Needs update: ' + (outstr ? 'Yes' : 'No');
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_ADSDataRaw':
                scriptColData.columns(key).value = JSON.stringify(item_props);
                break;

            case 'MExt_ADSDataFormatted':
                scriptColData.columns(key).value = JSON.stringify(item_props, null, "\t");
                break;

            case 'MExt_HelperContainer':
                outstr = 'Format: [' + item_props.container_format + '], Codec: [' + item_props.container_codec + '], EncApp: [' + item_props.container_enc_app + '], VCount: [' + item_props.video_count + '], ACount: [' + item_props.audio_count + '], TCount: [' + item_props.text_count + '], OCount: [' + item_props.others_count + '], Additional: [' + item_props.format_additional + '], Extra: [' + JSON.stringify(item_props.extra) + ']';
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_HelperVideoCodec':
                outstr = 'Format: [' + item_props.video_format + '], Codec: [' + item_props.video_codec + '], EncLibName: [' + item_props.video_enc_libname + '], Additional: [' + item_props.video_format_additional + '], Extra: [' + JSON.stringify(item_props.video_extra) + ']';
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_HelperAudioCodec':
                outstr = 'Format: [' + item_props.audio_format + '], Codec: [' + item_props.audio_codec + '], Version: [' + item_props.audio_format_version + '], Profile: [' + item_props.audio_format_profile + '], Settings Mode: [' + item_props.audio_format_set_mode + '], Additional: [' + item_props.audio_format_additional + '], Extra: [' + JSON.stringify(item_props.audio_extra) + ']';
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VideoCount':
                outstr = item_props.video_count || 0;
                scriptColData.columns(key).group = '# Video Streams: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_AudioCount':
                outstr = item_props.audio_count || 0;
                scriptColData.columns(key).group = '# Audio Streams: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_TextCount':
                outstr = item_props.text_count || 0;
                scriptColData.columns(key).group = '# Text/Subtitle Streams: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_OthersCount':
                outstr = item_props.others_count || 0;
                scriptColData.columns(key).group = '# Other Streams: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;


            case 'MExt_CleanedUpName':
            case 'MExt_VideoCodec':
                if (!item_props.video_count) {
                    scriptColData.columns(key).sort = 0;
                    // this grouping makes sense mainly in mixed-format directories
                    scriptColData.columns(key).group = 'Video Codec: ' + (item_props.audio_count ? '<Audio Only>' : '');
                    break;
                }
                // add alternative '' to the end if you concatenate multiple fields
                _tmp0 = item_props.container_format.toUpperCase();
                _tmp1 = (item_props.video_format + '-' + item_props.video_codec + '-' + item_props.video_enc_libname).toUpperCase() || '';	// note last ''
                _tmp2 = (item_props.video_format + '-' + item_props.video_codec).toUpperCase() || '';										// note last ''
                _tmp3 = (item_props.video_format + '-' + item_props.video_enc_libname).toUpperCase() || '';	// note last ''
                _tmp4 = (item_props.video_format).toUpperCase() || '';
                _tmp5 = (item_props.video_codec).toUpperCase() || '';

                var lcv = config.get('lookup_codecs');
                outstr = 	lcv[_tmp0+'-'+_tmp1] ||
							lcv[_tmp0+'-'+_tmp2] ||
							lcv[_tmp0+'-'+_tmp3] ||
							lcv[_tmp0+'-'+_tmp4] ||
							lcv[_tmp0+'-'+_tmp5] ||
							lcv[_tmp1] ||
							lcv[_tmp2] ||
							lcv[_tmp3] ||
							lcv[_tmp4] ||
							lcv[_tmp5];
                if (outstr && outstr.length === 2) {
                    outstr = config.get('codec_use_short_variant') === true ? outstr[1] : outstr[0];
                }
                outstr = outstr ||
							(
							    item_props.video_format +
								' (Fallback, adjust LOOKUP_CODECS)'
								// (item_props.video_format_additional ? ' (' + item_props.video_format_additional + ')' : '')
							);
                scriptColData.columns(key).group = 'Video Codec: ' + outstr;
                scriptColData.columns(key).value = outstr;
                _vcodec = outstr; // buffer for 'Clean Name'
                break;


            case 'MExt_CleanedUpName':
            case 'MExt_AudioCodec':
                if (!item_props.audio_count) {
                    scriptColData.columns(key).sort = 0;
                    // this grouping makes sense mainly in mixed-format directories
                    scriptColData.columns(key).group = 'Audio Codec: ' + (item_props.video_count ? '<Video Only>' : '');
                    break;
                }
                // add alternative '' to the end if you concatenate multiple fields
                _tmp0 = item_props.container_format.toUpperCase();
                _tmp1 = (item_props.audio_format + '-' + item_props.audio_codec + '-' + item_props.audio_format_version + '-' + item_props.audio_format_profile + '-' + item_props.audio_format_set_mode).toUpperCase() || '';	// note last ''
                _tmp2 = (item_props.audio_format + '-' + item_props.audio_codec + '-' + item_props.audio_format_version + '-' + item_props.audio_format_profile).toUpperCase() || '';	// note last ''
                _tmp3 = (item_props.audio_format + '-' + item_props.audio_codec + '-' + item_props.audio_format_version).toUpperCase() || ''; 											// note last ''
                _tmp4 = (item_props.audio_format + '-' + item_props.audio_codec).toUpperCase() || ''; 																					// note last ''
                _tmp5 = (item_props.audio_format).toUpperCase() || ''; 																													// note last ''
                _tmp6 = (item_props.audio_codec).toUpperCase() || '';
                _tmp7 = (item_props.audio_format + '-' + item_props.audio_format_profile).toUpperCase() || ''; 																			// note last ''


                logger.verbose(selected_item.name + "\t" + '_tmp0: ' + _tmp0 + ', _tmp1: ' + _tmp1 + ', _tmp2: ' + _tmp2 + ', _tmp3: ' + _tmp3 + ', _tmp4: ' + _tmp4 + ', _tmp5: ' + _tmp5 + ', _tmp6: ' + _tmp6 + ', _tmp7: ' + _tmp7);

                var lca = config.get('lookup_codecs');
                outstr = 	lca[_tmp0+'-'+_tmp1] ||
							lca[_tmp0+'-'+_tmp2] ||
							lca[_tmp0+'-'+_tmp3] ||
							lca[_tmp0+'-'+_tmp4] ||
							lca[_tmp0+'-'+_tmp5] ||
							lca[_tmp1] ||
							lca[_tmp2] ||
							lca[_tmp3] ||
							lca[_tmp4] ||
							lca[_tmp5] ||
							lca[_tmp6] ||
							lca[_tmp7];
                if (outstr && outstr.length === 2) {
                    outstr = config.get('codec_use_short_variant') === true ? outstr[1] : outstr[0];
                }

                outstr = outstr || (
                    item_props.audio_format +
								' (Fallback, adjust LOOKUP_CODECS)'
								// (item_props.audio_format_additional ? ' (' + item_props.audio_format_additional + ')' : '')
                );
                scriptColData.columns(key).group = outstr;
                // add additional info always to the end, some codecs like AAC, DSD use this heavily
                if (config.get('codec_append_addinfo')) {
                    outstr += (item_props.audio_format_additional ? ' (' + item_props.audio_format_additional + ')' : '');
                }
                scriptColData.columns(key).group = 'Audio Codec: ' + outstr;
                scriptColData.columns(key).value = outstr;
                _acodec = outstr; // buffer for 'Clean Name'
                break;

            case 'MExt_VideoBitrate':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                if (item_props.container_format === 'Flash Video' && item_props.video_format ===  'Sorenson Spark') {
                    item_props.video_bitrate = item_props.overall_bitrate - item_props.audio_bitrate;
                    logger.error(selected_item.name + ' -- FALLBACK 1 detected while calculating video bitrate: ' + item_props.video_bitrate);
                }
                outstr = Math.floor((item_props.video_bitrate || 0) / 1000);
                scriptColData.columns(key).sort = outstr;
                outstr = outstr ? outstr + ' kbps' : '';
                scriptColData.columns(key).group = 'Video Bitrate: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_AudioBitrate':
                if (!item_props.audio_count) { scriptColData.columns(key).sort = 0; break; }
                outstr = Math.floor((item_props.audio_bitrate || 0) / 1000);
                if (!outstr) {
                // try alternative methods: filesize or streamsize * 8 / duration
                // it will fail for raw AAC files, since they do not store any bitrate or duration info at all
                // but might work for other files
                    var duration = item_props.audio_duration || item_props.duration || selected_item.metadata.mp3songlength || 0;
                    if (duration) {
                        if (item_props.audio_stream_size) {
                            outstr = item_props.audio_stream_size * 8 / duration;
                        } else if (item_props.video_count === 0 && item_props.file_size && item_props.audio_count) {
                            outstr = item_props.file_size * 8 / duration;
                        }

                        outstr = Math.round(outstr / 1000);
                        if (selected_item.metadata && selected_item.metadata.video && selected_item.metadata.video.mp3bitrate) {
                        // this is only for audio-only .TS files, for which MediaInfo cannot parse the audio bitrate
                        // but if we have the overall bitrate, we can subtract the 'video bitrate', which are in fact for DVD menus, etc.
                            outstr = selected_item.metadata.video.mp3bitrate;
                            logger.normal(selected_item.name + ' -- FALLBACK 1 detected while calculating audio bitrate: ' + outstr);
                        }
                    }
                    if (!outstr && selected_item.metadata && selected_item.metadata.audio && selected_item.metadata.audio.mp3bitrate) {
                        outstr = selected_item.metadata.audio.mp3bitrate;
                        logger.normal(selected_item.name + ' -- FALLBACK 2 detected while calculating audio bitrate: ' + outstr);
                    }
                }
                scriptColData.columns(key).sort = outstr;
                outstr = outstr ? outstr + ' kbps' : '';
                scriptColData.columns(key).group = 'Audio Bitrate: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_TotalBitrate':
                outstr = Math.floor((item_props.overall_bitrate || 0) / 1000);
                scriptColData.columns(key).sort = outstr;
                outstr = outstr ? outstr + ' kbps' : '';
                scriptColData.columns(key).group = 'Total Bitrate: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;


            case 'MExt_TotalDuration':
            case 'MExt_VideoDuration':
            case 'MExt_AudioDuration':
            case 'MExt_CombinedDuration':
                var td = item_props.duration,
                    vd = item_props.video_duration,
                    ad = item_props.audio_duration;
                var tolerance = 5;
                var extended_info_on_mismatch = false;
                var grpPrefix;
                switch(key) {
                    case 'MExt_TotalDuration':		grpPrefix = 'Total '; break;
                    case 'MExt_VideoDuration':		grpPrefix = 'Video '; break;
                    case 'MExt_AudioDuration':		grpPrefix = 'Audio '; break;
                    case 'MExt_CombinedDuration':	grpPrefix = ''; break;
                }

                if (key === 'MExt_AudioDuration' && !item_props.audio_count) {
                    logger.normal('Col: ' + key + "\t" + selected_item.name + ' -- FALLBACK A1 - No audio');
                    scriptColData.columns(key).value = item_props.audio_bitrate ? '≠00:00' : '';
                    scriptColData.columns(key).sort  = scriptColData.columns(key).value || '00:00';
                    scriptColData.columns(key).group = grpPrefix + 'Duration: No Audio';
                    break;
                } else if (key === 'MExt_VideoDuration' && !item_props.video_count) {
                    logger.normal('Col: ' + key + "\t" + selected_item.name + ' -- FALLBACK V1 - No video');
                    scriptColData.columns(key).value = item_props.video_bitrate ? '≠00:00' : '';
                    scriptColData.columns(key).sort  = '00:00';
                    scriptColData.columns(key).group = grpPrefix + 'Duration: No Video';
                    break;
                } else if ((key === 'MExt_TotalDuration' || key === 'MExt_CombinedDuration') && !item_props.video_count && !item_props.audio_count) {
                    logger.force('Col: ' + key + "\t" + selected_item.name + ' -- This should have never happened!\nVideo count: ' + item_props.video_count + ', Audio count: ' + item_props.audio_count);
                    break;
                }

                if (key === 'MExt_AudioDuration' && ad) {
                    outstr = SecondsToHHmm(ad);
                } else if (key === 'MExt_VideoDuration' && vd) {
                    outstr = SecondsToHHmm(vd);
                } else if (key === 'MExt_TotalDuration' && td) {
                    outstr = SecondsToHHmm(td);
                } else if (key === 'MExt_CombinedDuration') {
                    outstr = SecondsToHHmm(td);
                }

                if (key === 'MExt_CombinedDuration') {
                    if (td && vd && ad) {
                    // all 3 exist
                        if (
                            (td === vd || (td >= vd && td-tolerance <= vd) || (td <= vd && td-tolerance >= vd) ) &&
							(td === ad || (td >= ad && td-tolerance <= ad) || (td <= ad && td-tolerance >= ad) ) &&
							(vd === ad || (td >= ad && td-tolerance <= ad) || (td <= ad && td-tolerance >= ad) )
                        ) {
                            logger.info(sprintf("All exist, all within tolerance limits -- td: %s, vd: %s, ad: %s", td, vd, ad));
                            outstr = SecondsToHHmm(td);
                        } else {
                            logger.info(sprintf("Col: %s\t%s All exist, some outside tolerance limits -- td: %s, vd: %s, ad: %s", key, selected_item.name, td, vd, ad));
                            outstr += (!extended_info_on_mismatch ? ' (!)' :
                                ' ('
										+ ( td ? 'T: '  + SecondsToHHmm(td) : '' )
										+ ( vd ? ' V: ' + SecondsToHHmm(vd) : '' )
										+ ( ad ? ' A: ' + SecondsToHHmm(ad) : '' )
										+ ')');
                        }
                    } else if (!td && !vd && !ad) {
                    // none of 3 exists - but if we are this far, we have at least 1 video or audio stream
                        outstr = '≠00:00';
                    } else {
                        logger.info(sprintf("Only some exist -- td: %s, vd: %s, ad: %s", td, vd, ad));
                        // check for tolerance
                        if 	(
                            (td && vd && td >= vd && td - tolerance <= vd) ||	// only slightly longer
								(td && ad && td >= ad && td - tolerance <= ad) ||	// only slightly longer
								(vd && ad && vd >= ad && vd - tolerance <= ad) ||	// only slightly longer
								(td && vd && td <= vd && td + tolerance >= vd) ||	// only slightly shorter
								(td && ad && td <= ad && td + tolerance >= ad) ||	// only slightly shorter
								(vd && ad && vd <= ad && vd + tolerance >= ad) 	// only slightly shorter
                        ) {
                        // within tolerance limit
                            logger.verbose(sprintf("Some within tolerance limits -- td: %s, vd: %s, ad: %s", td, vd, ad));
                            logger.verbose(sprintf("Some within tolerance limits -- %s <= %s <= %s", (td - tolerance), vd, (td + tolerance) ));
                            logger.verbose(sprintf("Some within tolerance limits -- %s <= %s <= %s", (td - tolerance), ad, (td + tolerance) ));
                            outstr = SecondsToHHmm(td);
                        } else {
                            logger.verbose(sprintf("Some outside tolerance limits -- td: %s, vd: %s, ad: %s", td, vd, ad));
                            logger.verbose(sprintf("Some outside tolerance limits -- %s <= %s <= %s", (td - tolerance), vd, (td + tolerance) ));
                            logger.verbose(sprintf("Some outside tolerance limits -- %s <= %s <= %s", (td - tolerance), ad, (td + tolerance) ));
                            outstr += (!extended_info_on_mismatch ? ' (!)' :
                                ' ('
										+ ( td ? 'T: '  + SecondsToHHmm(td) : '' )
										+ ( vd ? ' V: ' + SecondsToHHmm(vd) : '' )
										+ ( ad ? ' A: ' + SecondsToHHmm(ad) : '' )
										+ ')');
                        }
                    }
                }

                var ldg = config.get('lookup_duration_groups');
                if (ldg) {
                    for (var kd in ldg) {
                        if (item_props.audio_duration <= parseInt(kd)) {
                            scriptColData.columns(key).group = grpPrefix + 'Duration: ' + ldg[kd]; break;
                        }
                    }
                } else {
                    scriptColData.columns(key).group = grpPrefix + 'Duration: ' + outstr;
                }

                switch(key) {
                    case 'MExt_TotalDuration':
                        scriptColData.columns(key).sort  = outstr || (item_props.overall_bitrate || item_props.duration ? '00:01' : '00:00');
                        scriptColData.columns(key).value = outstr || (item_props.overall_bitrate || item_props.duration ? '≠00:00' : '');
                        break;
                    case 'MExt_VideoDuration':
                        scriptColData.columns(key).sort  = outstr || (item_props.video_bitrate || item_props.video_count ? '00:01' : '00:00');
                        scriptColData.columns(key).value = outstr || (item_props.video_bitrate || item_props.video_count ? '≠00:00' : '');
                        break;
                    case 'MExt_AudioDuration':
                        scriptColData.columns(key).sort  = outstr || (item_props.audio_bitrate || item_props.audio_count ? '00:01' : '00:00');
                        scriptColData.columns(key).value = outstr || (item_props.audio_bitrate || item_props.audio_count ? '≠00:00' : '');
                        break;
                    case 'MExt_CombinedDuration':
                        scriptColData.columns(key).sort  = outstr;
                        scriptColData.columns(key).value = outstr;
                        break;
                }
                break;


            case 'MExt_MultiAudio':
                outstr = (item_props.audio_count > 1 ? 'Yes' : 'No');
                scriptColData.columns(key).group = 'Has Multi-Track Audio: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_AudioChannels':
                outstr = item_props.audio_channels || '';
                if (!outstr) {
                    if (item_props.audio_count) {
                    // audio track does not report number of channels but file has audio! - only happens with Musepack & Raw DTS in my tests
                        scriptColData.columns(key).sort = 1;
                        outstr = 'X';
                    } else {
                        scriptColData.columns(key).sort = 0;
                        outstr = '0';
                    }
                }
                scriptColData.columns(key).group = 'Audio Channels: ' + (outstr || ''); // must come before next line, outstr is ' ' for empty!
                var lc = config.get('lookup_channels');
                if (lc) {
                    outstr = lc[outstr] ? lc[outstr] : outstr;
                    scriptColData.columns(key).group = 'Audio Channels: ' + (outstr === lc['0'] ? 'None' : outstr);
                }
                scriptColData.columns(key).value = outstr;
                break;


            case 'MExt_AudioLang':
                if (!item_props.audio_count) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.audio_language ? item_props.audio_language.toUpperCase() : '';
                scriptColData.columns(key).group = 'Audio Language: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VARDisplay':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.video_display_AR || '';
                scriptColData.columns(key).group = 'Display Aspect Ratio: ' + outstr;
                scriptColData.columns(key).sort = outstr || 0;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VARRaw':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                // outstr = item_props.video_width && item_props.video_height ? Math.toPrecision( item_props.video_width / item_props.video_height, 3) : 'n/a' ;
                // no toPrecision() support in JScript
                outstr = item_props.video_width && item_props.video_height ? Math.round(1000*item_props.video_width/item_props.video_height)/1000 : 'n/a' ;
                scriptColData.columns(key).group = 'Raw Aspect Ratio: ' + outstr;
                scriptColData.columns(key).sort = outstr || 0;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VARCombined':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                var _rawAR  = item_props.video_width && item_props.video_height ? Math.round(1000*item_props.video_width/item_props.video_height)/1000 : '';
                var _dispAR = item_props.video_display_AR || '';
                logger.verbose(selected_item.name + ', raw: ' + _rawAR + ', disp: ' + _dispAR);
                // outstr = _rawAR + (_rawAR != _dispAR ? ' (' + _dispAR + ')' : ''); // display
                outstr = _rawAR + (_rawAR != _dispAR ? ' (' + _dispAR + ')' : '');
                scriptColData.columns(key).group = 'Combined Aspect Ratio: ' + outstr;
                scriptColData.columns(key).sort  = _rawAR || 0;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VDimensions':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                // Variant 1: default to MediaInfo value
                outstr = (item_props.video_width && item_props.video_height ? item_props.video_width + ' x ' + item_props.video_height : selected_item.metadata.video.dimensions) || '' ;
                // Variant 2: default to DOpus value - not tested
                // outstr = selected_item.metadata.video.dimensions || (item_props.video_width && item_props.video_height ? item_props.video_width + ' x ' + item_props.video_height : '');
                scriptColData.columns(key).group = 'Dimensions: ' + outstr;
                scriptColData.columns(key).sort  = outstr || 0;
                scriptColData.columns(key).value = outstr;
                break;


            case 'MExt_CleanedUpName':
            case 'MExt_VResolution':
                if (!item_props.video_count || !item_props.video_width || !item_props.video_height) { scriptColData.columns(key).sort = 0; break; }
                var res_val = Math.min(item_props.video_height, item_props.video_width);
                var lr = config.get('lookup_resolutions');
                for (var kr in lr) {
                    if (res_val <= parseInt(kr)) {
                        outstr = lr[kr];
                        break; // for
                    }
                }
                if (item_props.video_height >= item_props.video_width && config.get('resolution_append_vertical')) {
                    outstr += ' (Vertical)';
                }
                scriptColData.columns(key).group = 'Resolution: ' + outstr;
                scriptColData.columns(key).value = outstr;
                _resolution = outstr; // buffer for 'Clean Name'
                break;

            case 'MExt_VFrameRate':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                // Variant 1: default to MediaInfo value
                outstr = (item_props.video_framerate ? item_props.video_framerate + ' fps': selected_item.metadata.video.framerate) || '' ;
                // Variant 2: default to DOpus value - not tested
                // outstr = selected_item.metadata.video.framerate || (item_props.video_framerate ? item_props.video_framerate + ' fps' : '');
                scriptColData.columns(key).group = 'Frame Rate: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VFrameCount':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.video_framecount ? item_props.video_framecount : '' ;
                scriptColData.columns(key).group = 'Frame Count: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_AudioBitrateMode':
                outstr = item_props.audio_bitrate_mode || '';
                scriptColData.columns(key).group = 'Audio Bitrate Mode: ' + (outstr || 'Unknown/Not Reported');
                scriptColData.columns(key).sort = outstr || 0;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_AudioCompressionMode':
                outstr = item_props.audio_compression_mode || '';
                scriptColData.columns(key).group = 'Audio Compression Mode: ' + (outstr || 'Unknown');
                scriptColData.columns(key).sort  = outstr || 0;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_HasReplayGain':
                outstr = item_props.has_replay_gain ? 'Yes' : '';
                scriptColData.columns(key).group = 'Has ReplayGain: ' + (outstr || 'No');
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_SubtitleLang':
                if (!item_props.text_count) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.text_language ? item_props.text_language.toUpperCase() : '';
                scriptColData.columns(key).group = 'Subtitle Language: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_GrossByterate':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                // this column shows the gross byterate (not bitrate) of the file, i.e. filesize/duration in KB
                // see next column below
                if (item_props.duration) {
                    outstr = Math.round( Math.round( selected_item.size / item_props.duration ) / 1024 );
                }
                scriptColData.columns(key).group = 'Gross Byterate: ' + outstr + ' kBps';
                scriptColData.columns(key).sort  = outstr || 0;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VBitratePerPixel':
                if (!item_props.video_count) { scriptColData.columns(key).sort = 0; break; }
                // this column shows the gross bits per pixel per second, i.e. video_size*8/(duration*width*height)
                // it is a very rough metric to show how 'efficient' the video compression of a file is
                // to me it does not make sense for audio files, because most users are well familiar with audio bitrates, often simply 2.0 44.1/48 kHz files
                //
                // this value does not need to be very high with higher res videos or very low with low res videos
                // ideally the higher the resolution, the higher bitrate should be
                // but if you use a lossless video codec or very high bitrate e.g. 3000 kbps for a tiny 320x240 video,
                // then this value will be accordingly very high
                // and similarly if you use e.g. a mere 3000 kbps for a 4K video, this value will be lower than it should;
                // typical bitrate/pixel: 1.5 to 7 for 4K movies in my samples, e.g. 10000 to 50000 kps
                // values like > 30 should be encountered only for lossless codecs like HuffYUV, not lossy codecs
                // personally I find a value above ~10-15 for this column is too high (possible reason: old/inefficient codec or bitrate wasted)
                // and a value below 1 is too low (possible reason: too heavily compressed, for streaming-only, blocking effects)
                // but some files with large static areas can be compressed very well and
                // newer codecs like HEVC/X265 are great at very low bitrates
                //
                // so there's no universal formula
                // just experiment and see or use this as a template for your own formula
                if (item_props.video_bitrate && item_props.video_width && item_props.video_height) {
                // round to 3 decimal digits
                    outstr = Math.round( 1000 * ( item_props.video_bitrate / ( item_props.video_width * item_props.video_height ) ) ) / 1000;
                }
                scriptColData.columns(key).group = 'Bitrate/Pixel: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VEncLibName':
                if (!item_props.video_enc_libname) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.video_enc_libname || '';
                scriptColData.columns(key).group = 'Encoded Library Name: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_VEncLib':
                if (!item_props.video_enc_lib) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.video_enc_lib || '';
                scriptColData.columns(key).group = 'Encoded Library: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;


            case 'MExt_VCodecID':
                if (!item_props.video_codec) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.video_codec || '';
                scriptColData.columns(key).group = 'Video Codec ID: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_ACodecID':
                if (!item_props.audio_codec) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.audio_codec || '';
                scriptColData.columns(key).group = 'Audio Codec ID: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_AFormatVersion':
                if (!item_props.audio_format_version) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.audio_format_version || '';
                scriptColData.columns(key).group = 'Audio Format Version: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_AProfile':
                if (!item_props.audio_format_profile) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.audio_format_profile || '';
                scriptColData.columns(key).group = 'Audio Format Profile: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_EncoderApp':
                if (!item_props.container_enc_app) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.container_enc_app || '';
                scriptColData.columns(key).group = 'Encoder App: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_DateEncoded':
                if (!item_props.container_date_encoded) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.container_date_encoded || '';
                scriptColData.columns(key).group = 'Date Encoded: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            case 'MExt_DateTagged':
                if (!item_props.container_date_tagged) { scriptColData.columns(key).sort = 0; break; }
                outstr = item_props.container_date_tagged || '';
                scriptColData.columns(key).group = 'Date Tagged: ' + outstr;
                scriptColData.columns(key).value = outstr;
                break;

            default:
<<<<<<< HEAD
                // check if there are any user-defined fields
=======
>>>>>>> b8370f1fc2d47fc98439d223376699b1629d4f27
                if(colExtraMap.exists(key)){
                    if (!item_props.extra[key]) { scriptColData.columns(key).sort = 0; break; }
                    outstr = item_props.extra[key] || '';
                    scriptColData.columns(key).group = colExtraMap.get(key) + ': ' + outstr;
                    scriptColData.columns(key).value = outstr;
                } else {
<<<<<<< HEAD
                    // nothing, default to empty string
                    outstr = '';
                }
=======
                    DOpus.output('nothing found');
                    outstr = '';
                }
                // nothing, default to empty string
>>>>>>> b8370f1fc2d47fc98439d223376699b1629d4f27
        } // switch

    } // for enum

    // // DOpus.output('name_cleanup: ' + JSON.stringify(config.get('name_cleanup'), null, 4));
    // if (config.get('name_cleanup')) {
    // 	key = 'MExt_CleanedUpName';
    // 	outstr = getCleanName(selected_item, _vcodec, _acodec, _resolution);
    // 	scriptColData.columns(key).group = 'Clean Name: ' + outstr;
    // 	scriptColData.columns(key).value = outstr;
    // }


    var ts2 = new Date().getTime();
    logger.verbose('OnMExt_MultiColRead() -- Elapsed: ' + (ts2 - ts1) + ', current: ' + ts2);
}

function getCleanName(oItem, sVCodec, sACodec, sResolution) {
    // DOpus.output('oItem.name_stem: ' + oItem.name_stem + ' - ' + sVCodec  + ' ' + sACodec + ' ' + sResolution);

    var reDateMatcher = new RegExp(/^(.+?)\s*(\d{4})/);

    var oldNameStem, oldExt, newNameStem, newExt;
    // If we're renaming a file then remove the extension from the end and save it for later.
    if (oItem.is_dir || oItem.is_junction || oItem.is_reparse || oItem.is_symlink) {
        return oItem.name;
    } else {
        oldNameStem = oItem.name_stem;
        oldExt      = oItem.ext;
    }

    var nameCleanupRegexes = config.get('name_cleanup'),
        arrNameOnly        = nameCleanupRegexes.nameOnly,
        arrExtOnly         = nameCleanupRegexes.extOnly;

    function massREReplace(str, arrRegexps) {
        var out = str;
        for (var i = 0; i < arrRegexps.length; i++) {
            var r = arrRegexps[i];
            out = out.replace(new RegExp(r[0], r[1]), r[2]);
        }
        return out;
    }

    newNameStem = massREReplace(oldNameStem, arrNameOnly);
    newExt      = massREReplace(oldExt,      arrExtOnly);

    // now that we have the names, find the year
    var year = newNameStem.match(reDateMatcher);
    if (year) {
        newNameStem = year[1] + ' - ' + year[2] + ' - ' + sVCodec + ' ' + sACodec + ' ' + sResolution;
        // get rid of extra dash's if the file was already renamed
        newNameStem = newNameStem.replace(/-\s*-/g, '-');
    }

    return newNameStem + newExt;

}

// UPDATE METADATA
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_Update(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;

    scriptCmdData.func.command.SetModifier("noprogress");

    var progress_bar = scriptCmdData.func.command.Progress;
    // progress_bar.skip  = true; // does not make sense, individual files are processed very fast already
    progress_bar.pause = true;
    progress_bar.abort = true;
    progress_bar.Init(scriptCmdData.func.sourcetab, 'Please wait'); 			// window title
    progress_bar.SetStatus('Updating metadata info in ADS via MediaInfo CLI'); 	// header
    progress_bar.Show();
    progress_bar.SetFiles(scriptCmdData.func.sourcetab.selected_files.count);
    progress_bar.Restart();

    if (config.get('force_refresh_after_update')) { var orig_selected_items = scriptCmdData.func.sourcetab.selected; }

    var selected_files_cnt = scriptCmdData.func.sourcetab.selstats.selfiles;
    logger.verbose("Selected files count: " + selected_files_cnt);

    fileloop: for (var e = new Enumerator(scriptCmdData.func.sourcetab.selected), cnt = 0; !e.atEnd(); e.moveNext(), cnt++) {
        var selected_item = e.item();
        if (selected_item.is_dir || selected_item.is_reparse || selected_item.is_junction || selected_item.is_symlink ) {
            logger.verbose('skipping ' + selected_item.name);
            continue;
        }
        logger.normal('Name: ' + selected_item.name);

        progress_bar.StepFiles(1);
        progress_bar.SetTitle(Global.SCRIPT_PREFIX + ' Update: ' + cnt + '/' + selected_files_cnt);
        progress_bar.SetName(selected_item.name);
        progress_bar.SetType('file');
        // progress_bar.EnableSkip(true, false);
        switch (progress_bar.GetAbortState()) {
        // case 's': progress_bar.ClearAbortState(); continue;
            case 'a': break fileloop;
            case 'p':
                while (progress_bar.GetAbortState() !== '') {
                    DOpus.Delay(200);
                    if (progress_bar.GetAbortState() === 'a') {
                        break fileloop;
                    }
                }
                break;
        }

        var json_contents = GetMediaInfoFor(selected_item);
        if (!json_contents) {
            logger.error('...no valid info received from MediaInfo');
            continue;
        }

        var mediainfo = JSON.parse(json_contents)['media'];
        logger.verbose('1st line: ' + mediainfo['@ref']);

        var out_obj = {
            'last_modify'	: new Date(selected_item.modify).valueOf(),
            'added_to_cache': 0 // this timestamp will be used to limit the number of items in cache
        };

        for (var i = 0; i < mediainfo['track'].length; i++) {
            var track = mediainfo['track'][i];
            logger.verbose('Track #' + i + ': ' + track['@type'] + ', ' + track['Format']);
            switch (track['@type']) {
                case 'General':
                    if (out_obj.container_format) break;	// read only 1st one
                    out_obj.container_format		= track['Format']							|| '';
                    out_obj.container_codec 		= track['CodecID']							|| '';
                    out_obj.container_enc_app		= track['Encoded_Application']				|| '';	// video does not have app, only container
                    out_obj.container_enc_lib		= track['Encoded_Library']					|| '';
                    out_obj.container_date_encoded	= track['Encoded_Date']				    	|| '';
                    out_obj.container_date_tagged	= track['Tagged_Date']				    	|| '';
                    out_obj.file_size 				= parseInt(track['FileSize'])				|| 0;
                    out_obj.duration				= parseFloat(track['Duration'])				|| 0;
                    out_obj.overall_bitrate			= parseInt(track['OverallBitRate'])			|| 0;
                    out_obj.overall_bitrate_mode	= track['OverallBitRate_Mode']				|| '';
                    out_obj.video_count 			= parseInt(track['VideoCount'])				|| 0;
                    out_obj.audio_count 			= parseInt(track['AudioCount'])				|| 0;
                    out_obj.text_count 				= parseInt(track['TextCount'])				|| 0;
                    out_obj.others_count 			= parseInt(track['MenuCount'])				|| 0;
                    out_obj.has_replay_gain			= (track['extra'] && (track['extra']['REPLAYGAIN_GAIN'] || track['extra']['replaygain_track_gain'] || track['extra']['replaygain_album_gain'] || track['extra']['R128_TRACK_GAIN'] || track['extra']['R128_ALBUM_GAIN'] )) ? 1 : 0;
                    out_obj.format_additional 		= track['Format_AdditionalFeatures']		|| '';
                    out_obj.extra					= track['extra'] 							|| '';

                    // if (track['extra']) {
                    // 	if (track['extra']['ENCODER_OPTIONS']) {
                    // 		// currently only for Opus files
                    // 		out_obj.audio_bitrate_mode 		= track['extra']['ENCODER_OPTIONS'].indexOf('--cvbr') >= 0 ? 'VBR' : out_obj.audio_bitrate_mode;
                    // 		out_obj.audio_bitrate_mode 		= track['extra']['ENCODER_OPTIONS'].indexOf('--cabr') >= 0 ? 'ABR' : out_obj.audio_bitrate_mode;
                    // 		out_obj.audio_bitrate_mode 		= track['extra']['ENCODER_OPTIONS'].indexOf('--ccbr') >= 0 ? 'CBR' : out_obj.audio_bitrate_mode;
                    // 		out_obj.audio_compression_mode 	= track['extra']['ENCODER_OPTIONS'].indexOf('--music') >= 0 ? 'Music' : out_obj.audio_compression_mode;
                    // 		out_obj.audio_compression_mode 	= track['extra']['ENCODER_OPTIONS'].indexOf('--speech') >= 0 ? 'Speech' : out_obj.audio_compression_mode;
                    // 	}
                    // 	// currently only for TS files - does not seem to be necessary and it's a very bad solution
                    // 	// if (out_obj.container_format === 'MPEG-TS' && track['extra']['OverallBitRate_Precision_Min'] && track['extra']['OverallBitRate_Precision_Max']) {
                    // 	// 	out_obj.audio_bitrate = Math.round( (track['extra']['OverallBitRate_Precision_Min'] + track['extra']['OverallBitRate_Precision_Max']) / 2*1000);
                    // 	// }
                    // 	// log.error("Min: " + track['extra']['OverallBitRate_Precision_Min'] + "\nMax: " + track['extra']['OverallBitRate_Precision_Max']);
                    // }
                    break;
                case 'Video':
                    if (out_obj.video_format) break;		// read only 1st video track
                    out_obj.video_format			= track['Format'] 							|| '';
                    out_obj.video_codec 			= track['CodecID']							|| '';
                    out_obj.video_duration			= parseFloat(track['Duration'])				|| out_obj.duration || 0;
                    out_obj.video_bitrate 			= parseInt(track['BitRate'])				|| parseInt(track['BitRate_Nominal']) || (track['extra'] && track['extra']['FromStats_BitRate']) || 0;
                    out_obj.video_width				= parseInt(track['Width'])					|| 0;
                    out_obj.video_height			= parseInt(track['Height'])					|| 0;
                    out_obj.video_framerate			= parseFloat(track['FrameRate'])			|| 0;
                    out_obj.video_framecount		= parseFloat(track['FrameCount'])			|| 0;
                    out_obj.video_display_AR		= parseFloat(track['DisplayAspectRatio'])	|| 0;
                    out_obj.video_compression_mode	= track['Compression_Mode']					|| '';
                    out_obj.video_format_additional = track['Format_AdditionalFeatures']		|| '';
                    out_obj.video_enc_lib			= track['Encoded_Library']					|| '';
                    out_obj.video_enc_libname		= track['Encoded_Library_Name']				|| '';
                    out_obj.video_stream_size		= parseInt(track['StreamSize'])				|| 0;
                    out_obj.video_extra				= track['extra'] 							|| '';
                    break;
                case 'Audio':
                    if (out_obj.audio_format) break;		// read only 1st audio track
                    out_obj.audio_format			= track['Format'] 							|| '';
                    out_obj.audio_codec 			= track['CodecID'] 							|| '';
                    out_obj.audio_language 			= track['Language'] 						|| 'und';
                    out_obj.audio_duration			= parseFloat(track['Duration'])				|| 0;
                    // out_obj.audio_bitrate 			= parseInt(track['BitRate'])				|| parseInt(track['extra'] && track['extra']['FromStats_BitRate']) || (out_obj.video_count === 0 ? out_obj.overall_bitrate : 0) || out_obj.audio_bitrate || 0;
                    out_obj.audio_bitrate 			= parseInt(track['BitRate'])				|| parseInt(track['extra'] && track['extra']['FromStats_BitRate']) || out_obj.audio_bitrate || 0;
                    out_obj.audio_format_version	= track['Format_Version'] 					|| '';
                    out_obj.audio_format_profile	= track['Format_Profile'] 					|| '';
                    out_obj.audio_format_set_mode	= track['Format_Settings_Mode'] 			|| '';
                    out_obj.audio_format_additional = track['Format_AdditionalFeatures']		|| '';
                    // out_obj.audio_bitrate_mode		= track['BitRate_Mode'] 					|| out_obj.audio_bitrate_mode || ( out_obj.audio_format === "Monkey's Audio" ? 'VBR' : '') || '';
                    out_obj.audio_bitrate_mode		= track['BitRate_Mode'] 					|| out_obj.audio_bitrate_mode || (config.get('formats_regex_vbr').test(out_obj.audio_format) && 'VBR') || '';
                    out_obj.audio_channels			= parseInt(track['Channels'])				|| 0;
                    out_obj.audio_bit_depth			= parseInt(track['BitDepth'])				|| 0;
                    out_obj.audio_sampling_rate		= parseInt(track['SamplingRate'])			|| 0;
                    // out_obj.audio_compression_mode  = track['Compression_Mode'] 				|| (config.get('formats_regex_lossless').test(out_obj.audio_format) && 'Lossless') || (config.get('formats_regex_lossy').test(out_obj.audio_format) && 'Lossy') || '';
                    // logger.error(out_obj.audio_format);
                    // logger.error(config.get('formats_regex_lossless'));
                    // logger.error(config.get('formats_regex_lossless').test(out_obj.audio_format));
                    out_obj.audio_compression_mode  = (config.get('formats_regex_lossless').test(out_obj.audio_format) && 'Lossless') || (config.get('formats_regex_lossy').test(out_obj.audio_format) && 'Lossy') || track['Compression_Mode']	|| '';
                    out_obj.audio_stream_size		= parseInt(track['StreamSize'])				|| 0;
                    out_obj.audio_extra				= track['extra'] 							|| '';

                    if (out_obj.audio_format === 'WMA' && track['Format_Profile'] === 'Lossless'){
                        out_obj.audio_compression_mode = 'Lossless';
                    }

                    out_obj.has_replay_gain 		= (out_obj.has_replay_gain || track['ReplayGain_Gain']) ? 1 : 0;	// RG can be stored on container level e.g. in M4A, MKA, or audio level, e.g. in MP3
                    break;
                case 'Text':
                    if (out_obj.text_codec) break;			// read only 1st subtitle track
                    out_obj.text_codec				= track['CodecID']  						|| '',
                    out_obj.text_title				= track['Title']	 						|| '';
                    out_obj.text_language			= track['Language'] 						|| '';
                    break;
                default:
                // add anything else than menus to the 'others'
                // optionally use track['extra'].length for a chapters count (may not be accurate, see MediaInfo dump first)
                    if (track['@type'] !== 'Menu') out_obj.others_count++;
            }
        }

        if (out_obj.audio_format === 'Opus') {
            // for Opus files
            if (out_obj.extra && out_obj.extra.ENCODER_OPTIONS) {
                logger.normal("Opus -- Encoder Options: " + out_obj.extra.ENCODER_OPTIONS);
                if (out_obj.extra.ENCODER_OPTIONS.indexOf('--cvbr') >= 0) {
                    out_obj.audio_bitrate_mode = 'VBR';
                } else if (out_obj.extra.ENCODER_OPTIONS.indexOf('--vbr') >= 0) {
                    out_obj.audio_bitrate_mode = 'VBR';
                } else if (out_obj.extra.ENCODER_OPTIONS.indexOf('--cabr') >= 0) {
                    out_obj.audio_bitrate_mode = 'VBR';
                } else if (out_obj.extra.ENCODER_OPTIONS.indexOf('--ccbr') >= 0) {
                    out_obj.audio_bitrate_mode = 'CBR';
                } else if (out_obj.extra.ENCODER_OPTIONS.indexOf('--hard-cbr') >= 0) {
                    out_obj.audio_bitrate_mode = 'CBR';
                }
                if (out_obj.extra.ENCODER_OPTIONS.indexOf('--music') >= 0) {
                    out_obj.audio_format_additional += 'Music';
                } else if (out_obj.extra.ENCODER_OPTIONS.indexOf('--speech') >= 0) {
                    out_obj.audio_format_additional += 'Speech';
                }
                logger.normal('Opus -- Determined BR mode: ' + out_obj.audio_bitrate_mode + "\nFormat Additional: " + out_obj.audio_format_additional);
            }

        }

        logger.verbose(JSON.stringify(out_obj));
        if (out_obj.video_count || out_obj.audio_count) {
            SaveMetadataADS(selected_item, out_obj);
        } else {
            logger.error(selected_item.name + ', skipping file - no video or audio stream found');
            continue fileloop;
        }
    }

    progress_bar.ClearAbortState();
    progress_bar.Hide();

    if (config.get('force_refresh_after_update')) {
        // refresh and re-select previously selected files
        // this seems to work for most of the time, but sometimes misses (?) few files if flat-folder view is used
        // not sure if it really misses files though
        util.cmdGlobal.RunCommand('Go Refresh');
        util.cmdGlobal.ClearFiles();
        for (var e = new Enumerator(orig_selected_items); !e.atEnd(); e.moveNext()) {
            util.cmdGlobal.AddFile(e.item());
        }
        util.cmdGlobal.RunCommand('Select FROMSCRIPT');
    }
}

// DELETE METADATA
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_Delete(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;

    scriptCmdData.func.command.SetModifier("noprogress");

    if (config.get('force_refresh_after_update')) { var orig_selected_items = scriptCmdData.func.sourcetab.selected; }

    var progress_bar = scriptCmdData.func.command.Progress;
    // progress_bar.skip  = true; // does not make sense, individual files are processed very fast already
    progress_bar.pause = true;
    progress_bar.abort = true;
    progress_bar.Init(scriptCmdData.func.sourcetab, 'Please wait'); // window title
    progress_bar.SetStatus('Deleting metadata ADS'); 				// header
    progress_bar.Show();
    progress_bar.SetFiles(scriptCmdData.func.sourcetab.selected.count);
    progress_bar.Restart();

    var selected_files_cnt = scriptCmdData.func.sourcetab.selstats.selfiles;
    logger.verbose("Selected files count: " + selected_files_cnt);

    fileloop: for (var e = new Enumerator(scriptCmdData.func.sourcetab.selected), cnt = 0; !e.atEnd(); e.moveNext(), cnt++) {
        var selected_item = e.item();
        if (selected_item.is_dir || selected_item.is_reparse || selected_item.is_junction || selected_item.is_symlink ) {
            logger.verbose('skipping ' + selected_item.name);
            continue;
        }
        logger.normal('Name: ' + selected_item.name);

        progress_bar.StepFiles(1);
        progress_bar.SetTitle(Global.SCRIPT_PREFIX + ' Delete: ' + cnt + '/' + selected_files_cnt);
        progress_bar.SetName(selected_item.name);
        progress_bar.SetType('file');
        // progress_bar.EnableSkip(true, false);
        switch (progress_bar.GetAbortState()) {
        // case 's': progress_bar.ClearAbortState(); continue;
            case 'a': break fileloop;
            case 'p':
                while (progress_bar.GetAbortState() !== '') {
                    DOpus.Delay(200);
                    if (progress_bar.GetAbortState() === 'a') {
                        break fileloop;
                    }
                }
                break;
        }

        DeleteMetadataADS(selected_item);
    }
    if (config.get('force_refresh_after_update')) {
        // refresh and re-select previously selected files
        // this seems to work for most of the time, but sometimes misses (?) few files if flat-folder view is used
        // not sure if it really misses files though
        util.cmdGlobal.RunCommand('Go Refresh');
        util.cmdGlobal.ClearFiles();
        for (var e = new Enumerator(orig_selected_items); !e.atEnd(); e.moveNext()) {
            util.cmdGlobal.AddFile(e.item());
        }
        util.cmdGlobal.RunCommand('Select FROMSCRIPT');
    }
    progress_bar.ClearAbortState();
    progress_bar.Hide();
    DumpCache('OnME_Delete', 2);
}

// CLEAR CACHE
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_ClearCache() {
    util.sv.set('cache', DOpus.Create.Map());
    logger.force('Cache cleared');
}

// DUMP ADS METADATA
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_ADSDump(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;
    DOpus.Output("\n\n\n\n\n");
    logger.force(GetMetadataForAllSelectedFiles(scriptCmdData));
    DOpus.Output("\n\n\n\n\n");
}

// COPY ADS METADATA
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_ADSCopy(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;

    util.cmdGlobal.RunCommand('Clipboard SET ' + GetMetadataForAllSelectedFiles(scriptCmdData));
    logger.force('Copied ADS metadata for selected files to clipboard');
}

// DUMP MEDIAINFO OUTPUT
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_MediaInfoDump(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;
    DOpus.Output("\n\n\n\n\n");
    logger.force(GetMediaInfoForAllSelectedFiles(scriptCmdData));
    DOpus.Output("\n\n\n\n\n");
}

// COPY MEDIAINFO OUTPUT
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_MediaInfoCopy(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;

    util.cmdGlobal.RunCommand('Clipboard SET ' + GetMediaInfoForAllSelectedFiles(scriptCmdData));
    logger.force('Copied MediaInfo output for selected files to clipboard');
}

// ESTIMATE BITRATES
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_EstimateBitrates(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;

    var target_bitsperpixel_values = [
        0.5, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 20, 25, 30, 50
    ];
    var out_obj = {};
    for (var e = new Enumerator(scriptCmdData.func.sourcetab.selected); !e.atEnd(); e.moveNext()) {
        var selected_item = e.item();
        if (selected_item.is_dir || selected_item.is_reparse || selected_item.is_junction || selected_item.is_symlink ) {
            logger.verbose('skipping ' + selected_item.name);
            continue;
        }
        var item_props = ReadMetadataADS(selected_item);
        out_obj[selected_item.realpath] = [];
        for (var i = 0; i < target_bitsperpixel_values.length; i++) {
            var current_target_value = target_bitsperpixel_values[i];
            var estimated_target_bitrate = Math.round(current_target_value * item_props.video_width * item_props.video_height / (1024)) + ' kbps';
            out_obj[selected_item.realpath].push( { 'target_bitsperpixel': current_target_value, 'estimated kbps': estimated_target_bitrate } );
        }
    }
    DOpus.Output(JSON.stringify(out_obj, null, 4));
}

// TOGGLE ESSENTIAL COLUMNS
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_ToggleEssentialColumns(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;
    _toggleColumnGroup('essential', config.get('fields_essential'), config.get('fields_essential_after'));
}
// TOGGLE OPTIONAL COLUMNS
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_ToggleOptionalColumns(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;
    _toggleColumnGroup('optional', config.get('fields_optional'), config.get('fields_optional_after'));
}
// TOGGLE OTHER COLUMNS
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_ToggleOtherColumns(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;
    _toggleColumnGroup('other', config.get('fields_other'), config.get('fields_other_after'));
}
// TOGGLE VERBOSE COLUMNS
// called by user command, button, etc.
// eslint-disable-next-line no-unused-vars
function OnME_ToggleVerboseColumns(scriptCmdData) {
    if (!validateConfigAndShowResult(false, scriptCmdData.func.Dlg, true)) return;
    _toggleColumnGroup('verbose', config.get('fields_verbose'), config.get('fields_verbose_after'));
}
// internal method called by OnME_ToggleXXXColumns
function _toggleColumnGroup(groupName, columnsArray, columnAfter) {
    var col_after = columnAfter.indexOf('MExt_') === 0 ? 'scp:' + Global.SCRIPT_NAME + '/' + columnAfter : columnAfter;
    var refresh;
    util.cmdGlobal.clearFiles();
    for (var i = 0; i < columnsArray.length; i++) {
        var item = columnsArray[i];
        refresh = false;
        if (item.indexOf('MExt_') >= 0) {
            item = 'scp:' + Global.SCRIPT_NAME + '/' + item;
            refresh = true;
        }
        logger.info('Toggling col: ' + item);
        cmd = 'Set COLUMNSTOGGLE=' + item + '(!' + (i+1) + '+' +  col_after + ',*)'; // * is for auto-size
        logger.verbose('_toggleColumnGroup -- (' + groupName + '):   ' + cmd);
        util.cmdGlobal.addLine(cmd);
    }
    util.cmdGlobal.run();
    util.cmdGlobal.clear();
    if (refresh) {
        // TODO - recheck if this is necessary
        // Script.RefreshColumn(columnsArray[i]);
    }
}
// VALIDATE CONFIG
// eslint-disable-next-line no-unused-vars
function OnME_ConfigValidate(scriptCmdData) {
    var debug = scriptCmdData.func.args.got_arg.DEBUG;
    var showValues = scriptCmdData.func.args.got_arg.SHOWVALUES;
    validateConfigAndShowResult(debug, showValues, scriptCmdData.func.Dlg, false);
}

// internal wrapper method called by most methods
// optionally shows a user dialog if configuration is invalid
function validateConfigAndShowResult(debug, showValues, dialog, skipIfValid, validateDefaultValues) {
    var isValid = validateConfig(debug, showValues, validateDefaultValues);
    if (!skipIfValid) {
        var dlgConfirm		= dialog;
        dlgConfirm.message  = Global.SCRIPT_NAME + ' configuration is ' + (isValid ? 'valid' : 'invalid');
        dlgConfirm.title	= Global.SCRIPT_NAME + ' - Configuration Validation' ;
        dlgConfirm.buttons	= 'OK';
        ret = dlgConfirm.show;
    }
    return isValid;
}

// internal method to validate the config - optionally validates script default values as well
function validateConfig(debug, showValues, validateDefaultValues) {
    if (debug) {
        logger.force('validateConfig started - debug: ' + debug + ', showValues: ' + showValues);
    } else {
        logger.normal('validateConfig started - debug: ' + debug + ', showValues: ' + showValues);
    }

    config.setErrorMode(config.modes.DIALOG);

    var configIsValid = true;
    var alternativeValue = 'value not shown for brevity';

    var cfgkeys = config.getKeys();

    if(validateDefaultValues) {
        if(debug) {
            logger.force('Number of items: ' + config.getCount() + ', matches actual keys count: ' + (config.getCount() == cfgkeys.length));
            logger.force('');
            logger.force('Dumping current config as string');
        }
        for (var i = 0; i < cfgkeys.length; i++) {
            var k = cfgkeys[i],
                t = config.getType(k),
                b = config.getBinding(k),
                v = config.get(k, false),
                iv= config.validate(v, t);
            configIsValid = configIsValid && iv;
            if (debug) {
                logger.force(
                    sprintf(
                        '#%02d key: %-30s type: %-15s bindTo: %-40s valid: %s  --  value: %s\n',
                        i+1,
                        k,
                        t,
                        b,
                        (iv ? iv : iv + ', expected: ' + t + ', found: ' + typeof v),
                        (!!showValues ? v : alternativeValue)
                    )
                );
            }
        }
    }


    if (debug) {
        logger.force('');
        logger.force('');
        logger.force('');
        logger.force('Checking bound values');
    }
    for (var i = 0; i < cfgkeys.length; i++) {
        var k = cfgkeys[i],
            t = config.getType(k),
            b = config.getBinding(k),
            v = config.get(k, true),
            iv = config.validate(v, t);
        configIsValid = configIsValid && iv;
        if (debug) {
            logger.force(
                sprintf(
                    '#%02d key: %-30s type: %-15s boundTo: %-40s valid: %s  --  bound value: %s\n',
                    i+1,
                    k,
                    t,
                    b,
                    (iv ? iv : iv + ', expected: ' + t + ', found: ' + typeof v + ' (boolean false is standard for invalid values)'),
                    (!!showValues ? v : alternativeValue)
                )
            );
        }
    }
    if (debug) {
        logger.force('');
        logger.force('');
        logger.force('');
        logger.force('validateConfig finished - config valid: ' + configIsValid);
    } else {
        logger.normal('validateConfig finished - config valid: ' + configIsValid);
    }

    return configIsValid;
}

// DEVELOPER HELPER METHOD
// eslint-disable-next-line no-unused-vars
function OnME_TestMethod1(scriptCmdData) {
    // unit tests for the poor
    var oldErrorMode = config.getErrorMode();

    DOpus.Output('OnME_TestMethod1 -- Adding valid values - started');
    config.setErrorMode(config.modes.ERROR);
    config.addBoolean('a', false);										DOpus.Output('\ttype: ' + config.getType('a'));
    config.addString('b', "hello world");								DOpus.Output('\ttype: ' + config.getType('b'));
    config.addNumber('c', 1);											DOpus.Output('\ttype: ' + config.getType('c'));
    config.addPath('d', '%gvdTool%\\MMedia\\MediaInfo\\MediaInfo.exe');	DOpus.Output('\ttype: ' + config.getType('d'));
    config.addArray('e', [ 'a', 1 ]);									DOpus.Output('\ttype: ' + config.getType('e'));
    config.addPOJO('f', { 'a': 1 });									DOpus.Output('\ttype: ' + config.getType('f'));
    config.addObject('g', { 'a': 1, fn: function(){} });				DOpus.Output('\ttype: ' + config.getType('g'));
    config.addRegexp('h', new RegExp(/foo.+bar/));						DOpus.Output('\ttype: ' + config.getType('h'));
    config.addJSON('i', '{ "a": 1 }');									DOpus.Output('\ttype: ' + config.getType('i'));
    config.addFunction('j', function(){});								DOpus.Output('\ttype: ' + config.getType('j'));
    config.setErrorMode(oldErrorMode);
    DOpus.Output('OnME_TestMethod1 -- Adding valid values - finished');

    DOpus.Output('');
    DOpus.Output('');
    DOpus.Output('');
    DOpus.Output('');

    DOpus.Output('OnME_TestMethod1 -- Adding invalid values - started');
    config.setErrorMode(config.modes.NONE);
    var res;
    res = config.addJSON('a', "{ 'a': 1 }");							DOpus.Output('\tresult: ' + res); // causes an error, JSON may not use single quotes
    res = config.addNumber('b', false);									DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addNumber('c', "hello world");							DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addString('d', 1);										DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addPath('e', 'MediaInfo.exe');							DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addJSON('f', [ 'a', 1 ]);								DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addArray('g', { 'a': 1 });								DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addPOJO('h', {"a":1,fn:function(){}});					DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addObject('i', '{ "a": 1');							DOpus.Output('\tresult: ' + res); // value not matching given type
    res = config.addString('j', function(){var x= 1});					DOpus.Output('\tresult: ' + res); // value not matching given type
    config.setErrorMode(oldErrorMode);
    DOpus.Output('OnME_TestMethod1 -- Adding invalid values - finished');

    DOpus.Output('');
    DOpus.Output('');
    DOpus.Output('');
    DOpus.Output('');



    DOpus.Output('OnME_TestMethod1 -- Testing error modes - started');
    var oldLoggerLevel = logger.getLevel();

    logger.force("Current error level: " + logger.getLevel());
    logger.error("This is a sample message in level: "   + "ERROR");
    logger.warn("This is a sample message in level: "    + "WARN");
    logger.normal("This is a sample message in level: "  + "NORMAL");
    logger.info("This is a sample message in level: "    + "INFO");
    logger.verbose("This is a sample message in level: " + "VERBOSE");

    DOpus.Output('');

    logger.setLevel(logger.levels.NONE);
    logger.force("Current error level: " + logger.getLevel());
    logger.error("This is a sample message in level: "   + "ERROR");
    logger.warn("This is a sample message in level: "    + "WARN");
    logger.normal("This is a sample message in level: "  + "NORMAL");
    logger.info("This is a sample message in level: "    + "INFO");
    logger.verbose("This is a sample message in level: " + "VERBOSE");

    DOpus.Output('');

    logger.setLevel(logger.levels.ERROR);
    logger.force("Current error level: " + logger.getLevel());
    logger.error("This is a sample message in level: "   + "ERROR");
    logger.warn("This is a sample message in level: "    + "WARN");
    logger.normal("This is a sample message in level: "  + "NORMAL");
    logger.info("This is a sample message in level: "    + "INFO");
    logger.verbose("This is a sample message in level: " + "VERBOSE");

    DOpus.Output('');

    logger.setLevel(logger.levels.WARN);
    logger.force("Current error level: " + logger.getLevel());
    logger.error("This is a sample message in level: "   + "ERROR");
    logger.warn("This is a sample message in level: "    + "WARN");
    logger.normal("This is a sample message in level: "  + "NORMAL");
    logger.info("This is a sample message in level: "    + "INFO");
    logger.verbose("This is a sample message in level: " + "VERBOSE");

    DOpus.Output('');

    logger.setLevel(logger.levels.NORMAL);
    logger.force("Current error level: " + logger.getLevel());
    logger.error("This is a sample message in level: "   + "ERROR");
    logger.warn("This is a sample message in level: "    + "WARN");
    logger.normal("This is a sample message in level: "  + "NORMAL");
    logger.info("This is a sample message in level: "    + "INFO");
    logger.verbose("This is a sample message in level: " + "VERBOSE");

    DOpus.Output('');

    logger.setLevel(logger.levels.INFO);
    logger.force("Current error level: " + logger.getLevel());
    logger.error("This is a sample message in level: "   + "ERROR");
    logger.warn("This is a sample message in level: "    + "WARN");
    logger.normal("This is a sample message in level: "  + "NORMAL");
    logger.info("This is a sample message in level: "    + "INFO");
    logger.verbose("This is a sample message in level: " + "VERBOSE");

    DOpus.Output('');

    logger.setLevel(logger.levels.VERBOSE);
    logger.force("Current error level: " + logger.getLevel());
    logger.error("This is a sample message in level: "   + "ERROR");
    logger.warn("This is a sample message in level: "    + "WARN");
    logger.normal("This is a sample message in level: "  + "NORMAL");
    logger.info("This is a sample message in level: "    + "INFO");
    logger.verbose("This is a sample message in level: " + "VERBOSE");

    logger.setLevel(oldLoggerLevel);
    DOpus.Output('OnME_TestMethod1 -- Testing error modes - finished');
}
// DEVELOPER HELPER METHOD
// eslint-disable-next-line no-unused-vars
function OnME_TestMethod2(scriptCmdData) {
    DOpus.ClearOutput();
    // validateConfigAndShowResult(true, scriptCmdData.func.Dlg, false, true);
}






/**********************************************************************************************************************
 *
 *
 *
 *  INTERNAL METHODS
 *
 *
 *
 **********************************************************************************************************************/




// internal method
// returns metadata for all selected files, calls ReadMetadataADS()
function GetMetadataForAllSelectedFiles(scriptCmdData) {
    var out_obj = {};
    var current_pojo = '';
    for (var e = new Enumerator(scriptCmdData.func.sourcetab.selected); !e.atEnd(); e.moveNext()) {
        var selected_item = e.item();
        if (selected_item.is_dir || selected_item.is_reparse || selected_item.is_junction || selected_item.is_symlink ) {
            logger.warn(selected_item.name + ' ...skipping');
            continue;
        }
        current_pojo = ReadMetadataADS(selected_item);
        if (!current_pojo || !isObject(current_pojo)) {
            logger.warn(selected_item.name + ' ...empty or invalid metadata, update first');
            continue;
        }
        out_obj[selected_item.realpath] = current_pojo;
    }
    return JSON.stringify(out_obj, null, 4);
}

// internal method
// reads raw MediaInfo output for all selected files, calls GetMediaInfoFor()
function GetMediaInfoForAllSelectedFiles(scriptCmdData) {
    var out_obj = {};
    for (var e = new Enumerator(scriptCmdData.func.sourcetab.selected); !e.atEnd(); e.moveNext()) {
        var selected_item = e.item();
        if (selected_item.is_dir || selected_item.is_reparse || selected_item.is_junction || selected_item.is_symlink ) {
            logger.verbose('skipping ' + selected_item.name);
            continue;
        }
        logger.normal('Name: ' + selected_item.name);
        current_json = JSON.parse(GetMediaInfoFor(selected_item));
        out_obj[selected_item.realpath] = current_json;
    }
    return JSON.stringify(out_obj, null, 4);
}

function GetDirectRunnerPath() {
    var sDirectRunnerPath = config.get('temp_files_dir') + '\\' + 'DOpusMExtScript-MediaInfoRunner.js';
    return sDirectRunnerPath;
}

function createDirectRunner(sTargetPath) {
    var sContents =
	'/*\n' +
	'https://github.com/cy-gh/CuDirectHiddenRunner\n' +
	'This file is automatically created by CuMediaExternder DOpus script\n' +
	'in order to get MediaInfo output without using temporary files.\n' +
	'*/\n' +
	'var oScriptArgs     = WScript.Arguments;\n' +
	'if (oScriptArgs.length !== 3) WScript.Quit(1);\n' +
	'var sPrefix         = oScriptArgs(0),\n' +
	'    iChunkSize      = parseInt(oScriptArgs(1), 10),\n' +
	'    sCommand        = oScriptArgs(2),\n' +
	'    oWSH            = WScript.CreateObject("WScript.Shell"),\n' +
	'    oEnvVolatile    = oWSH.Environment("Volatile");\n' +
	'if (WScript.FullName.toLowerCase().indexOf(\'wscript\') >= 0) {\n' +
	'    // rerun this script with cscript and pass the command we got - 0: hidden, true: wait\n' +
	'    oWSH.Run(\'cscript.exe //nologo "\' + WScript.ScriptFullName + \'" "\' + sPrefix + \'" "\' + iChunkSize + \'" "\' + sCommand + \'"\', 0, true);\n' +
	'} else {\n' +
	'    var sStdOut = \'\',\n' +
	'        sStdErr = \'\',\n' +
	'        iStdOutLines = 0,\n' +
	'        iStdErrLines = 0,\n' +
	'        iPtr    = 0,\n' +
	'        sChunk  = \'\';\n' +
	'    sCommand = \'"\' + sCommand.replace(/\'\'/g, \'"\') + \'"\';\n' +
	'    var oExec = oWSH.Exec(\'%comspec% /c \' + sCommand);\n' +
	'    do {\n' +
	'        while (!oExec.StdOut.AtEndOfStream) sStdOut += oExec.StdOut.ReadLine() + \'\\n\';\n' +
	'        while (!oExec.StdErr.AtEndOfStream) sStdErr += oExec.StdErr.ReadLine() + \'\\n\';\n' +
	'    } while (0 == oExec.Status && !oExec.AtEndOfStream);\n' +
	'    oEnvVolatile(sPrefix + \'EC\')    = oExec.ExitCode;\n' +
	'    if (sStdOut.length === 0) {\n' +
	'        oEnvVolatile(sPrefix + \'OUT0\')  = 0;\n' +
	'    } else if (sStdOut.length > iChunkSize) {\n' +
	'        // split\n' +
	'        for (iPtr = 0; iPtr < sStdOut.length; iStdOutLines++, iPtr += iChunkSize) {\n' +
	'            sChunk = sStdOut.slice(iPtr, iPtr + iChunkSize);\n' +
	'            oEnvVolatile(sPrefix + \'OUT\' + (iStdOutLines + 1)) = sChunk;\n' +
	'        }\n' +
	'        oEnvVolatile(sPrefix + \'OUT0\')  = iStdOutLines;\n' +
	'    } else {\n' +
	'        oEnvVolatile(sPrefix + \'OUT0\')  = 1;\n' +
	'        oEnvVolatile(sPrefix + \'OUT1\')  = sStdOut;\n' +
	'    }\n' +
	'    if (sStdErr.length === 0) {\n' +
	'        oEnvVolatile(sPrefix + \'ERR0\')  = 0;\n' +
	'    } else if (sStdErr.length > iChunkSize) {\n' +
	'        // split\n' +
	'        for (iPtr = 0; iPtr < sStdErr.length; iStdErrLines++, iPtr += iChunkSize) {\n' +
	'            sChunk = sStdErr.slice(iPtr, iPtr + iChunkSize);\n' +
	'            oEnvVolatile(sPrefix + \'ERR\' + (iStdErrLines + 1)) = sChunk;\n' +
	'        }\n' +
	'        oEnvVolatile(sPrefix + \'ERR0\')  = iStdErrLines;\n' +
	'    } else {\n' +
	'        oEnvVolatile(sPrefix + \'ERR0\')  = 1;\n' +
	'        oEnvVolatile(sPrefix + \'ERR1\')  = sStdErr;\n' +
	'    }\n' +
	'}\n'
    SaveFile(sTargetPath, sContents);
}

function getUniqueIDForString(sString) {
    var blob = DOpus.create().blob();
    blob.copyFrom(sString);
    return DOpus.fsUtil().hash(blob, 'md5');
}


// internal method
// reads raw MediaInfo output for a single file, always reads from disk
function GetMediaInfoFor(oItem) {

	function removeVolVar(sVar) {
		util.shellGlobal.Environment("Volatile").remove(sVar);
	}
    var mediainfo_output;
    var exit_code;
    var cmd;
    var sVarName;

    if (config.get('templess_mode')) {
        logger.normal('Running in TEMPLESS MODE');
        var iChunkSize = config.get('templess_chunk_size') || 32768;

        // check the hidden runner helper file which gets command output without temporary files
        var sDirectRunnerPath = GetDirectRunnerPath();
        if (!DOpus.fsUtil().exists(sDirectRunnerPath)) {
            createDirectRunner(sDirectRunnerPath); // create the file
        }
        // get a unique ID for this file
        var sPrefix = getUniqueIDForString('' + oItem.realpath);
        // run the helper
        cmd = sprintf(
            'wscript.exe "%s" "%s" "%s" "\'\'%s\'\' --Output=JSON \'\'%s\'\'"',
            sDirectRunnerPath,
            sPrefix,
            iChunkSize,
            config.get('mediainfo_path'),
            oItem.realpath
        );
        logger.normal('Running command: ' + cmd);
        exit_code = util.shellGlobal.Run(cmd, 0, true); // 0: hidden, true: wait

        // get the output environment variables
        var sExitCode = DOpus.fsUtil().resolve('%' + sPrefix + 'EC%'),
		    iStdOutChunks   = parseInt(DOpus.fsUtil().resolve('%' + sPrefix + 'OUT0%'), 10),
		    iStdErrChunks   = parseInt(DOpus.fsUtil().resolve('%' + sPrefix + 'ERR0%'), 10);

        logger.verbose(sprintf(
            // '\nHelper path: %s\nPrefix: %s\nExit Code: %d\nStdErr: %s\nStdOut: %s',
            '\nHelper path: %s\nPrefix: %s\nExit Code: %d\nSTDOUT Chunks: %d\nSTDERR Chunks: %d',
            sDirectRunnerPath,
            sPrefix,
            sExitCode,
            iStdOutChunks,
            iStdErrChunks
        ));

        if (!sExitCode || parseInt(sExitCode, 10) !== 0) {
            logger.error(oItem.name + ', error occurred while executing MediaInfo, exit code: ' + sExitCode);
            // TODO - join STDERR
            // var sStdErr = '';
            // // note the error chunks start at 1
            // for(var i = 1; i <= iStdErrChunks; i++) {
            // 	// get the var
            // 	sVarName = sPrefix + 'ERR' + i;
            // 	sStdErr += '' + DOpus.fsUtil().resolve('%' + sVarName + '%');
            // 	// delete the var
			// removeVolVar(sVarName);
            // }
            // logger.error(oItem.name + ', ' + sStdErr);
            return;
        }
        mediainfo_output = '';
        // note the output chunks start at 1
        for(var i = 1; i <= iStdOutChunks; i++) {
            // get the var
            sVarName = sPrefix + 'OUT' + i;
            mediainfo_output += '' + DOpus.fsUtil().resolve('%' + sVarName + '%');
            // delete the var
            removeVolVar(sVarName);
        }
        removeVolVar(sPrefix + 'EC');
        removeVolVar(sPrefix + 'OUT0');
        removeVolVar(sPrefix + 'ERR0');
        logger.normal('...finished');

    } else {

        logger.normal('Running in TEMPORARY FILES MODE');

        var json_filepath 	= config.get('temp_files_dir') + '\\' + oItem.name_stem + '.json';
        cmd 			= '"' + config.get('mediainfo_path') + '" --Output=JSON --LogFile="' + json_filepath + '" "' + oItem.realpath + '"';
        logger.verbose('Executing: ' + cmd);
        /*
			Run  can return %ERRORLEVEL%     but does NOT support async output, does not call CMD to execute the action
			Exec can NOT return %ERRORLEVEL% but supports real-time output to the user (if it is a command-line program)

			strErrorCode = objShell.Run(ipconfig, 0, true)

			comspec = objShell.ExpandEnvironmentStrings("%comspec%")
			Set objExec = objShell.Exec(comspec & " /c ipconfig")
			Do
				WScript.StdOut.WriteLine(objExec.StdOut.ReadLine())
			Loop While Not objExec.Stdout.atEndOfStream
			WScript.StdOut.WriteLine(objExec.StdOut.ReadAll)
		*/
        exit_code = util.shellGlobal.Run(cmd, 0, true); // 0: hidden, true: wait
        if (exit_code) {
            logger.error(oItem.name + ', error occurred while executing MediaInfo, exit code: ' + exit_code);
            return;
        }
        mediainfo_output = ReadFile(json_filepath, 'utf-8', false);
        util.cmdGlobal.RunCommand('Delete /quiet /norecycle "' + json_filepath + '"');
    }
    return mediainfo_output;
}

/*
	var res, str, enc;

	str = encodeURI(new Date().getTime().toString());
	enc = encoding.utf16;
	res = SaveFile("Y:\\MediaInfo\\victim.txt:SecondStream", str, enc);
	res = ReadFile("Y:\\MediaInfo\\victim.txt:SecondStream", enc);
	// log.force("\nOrig: " + str + "\nNew : " + res + "\n" + (res === str ? 'SUCCESS!' : ":("));
	log.force((res === str ? 'SUCCESS!' : ":("));

	str = encodeURI("{\"a\": 1}");
	enc = encoding.utf16;
	res = SaveFile("Y:\\MediaInfo\\victim.txt:SecondStream", str, enc);
	res = ReadFile("Y:\\MediaInfo\\victim.txt:SecondStream", enc);
	// log.force("\nOrig: " + str + "\nNew : " + res + "\n" + (res === str ? 'SUCCESS!' : ":("));
	log.force((res === str ? 'SUCCESS!' : ":("));

	str = encodeURI(new Date().getTime().toString());
	enc = encoding.utf8;
	res = SaveFile("Y:\\MediaInfo\\victim.txt:SecondStream", str, enc);
	res = ReadFile("Y:\\MediaInfo\\victim.txt:SecondStream", enc);
	// log.force("\nOrig: " + str + "\nNew : " + res + "\n" + (res === str ? 'SUCCESS!' : ":("));
	log.force((res === str ? 'SUCCESS!' : ":("));

	str = encodeURI("{\"a\": 1}");
	enc = encoding.utf8;
	res = SaveFile("Y:\\MediaInfo\\victim.txt:SecondStream", str, enc);
	res = ReadFile("Y:\\MediaInfo\\victim.txt:SecondStream", enc);
	// log.force("\nOrig: " + str + "\nNew : " + res + "\n" + (res === str ? 'SUCCESS!' : ":("));
	log.force((res === str ? 'SUCCESS!' : ":("));
	return;
*/

// internal method
// reads requested file (incl. ADS streams)
function ReadFile(path, format) {
    // path, use "Y:\\Path\\file.txt" or "Y:\\Path\\file.txt:CustomMetaInfo" for ADS
    // format: "base64", "quoted", "auto"=not supplied, "utf-8", "utf-16", "utf-16-le", "utf-16-be"
    // the only ones which worked reliably in my tests are utf-8 & utf-16, since they're the only ones Blob.CopyFrom() supports
    // use only one of encoding.utf8 or encoding.utf16
    //
    // res = ReadFile("Y:\\MediaInfo\\victim.txt:SecondStream", encoding.utf8);
    // res = ReadFile("Y:\\MediaInfo\\victim.txt", encoding.utf16);
    var res = false;
    format = format === TEXT_ENCODING.utf16 ? 'utf-16' : 'utf-8';

    if (!DOpus.FSUtil.Exists(path)) { return res; }

    var fh = DOpus.FSUtil.OpenFile(path); // default read mode
    if(fh.error !== 0) {
        // if you enable this log level in the global options you will receive a lot of messages for missing streams of unprocessed files
        logger.normal(path + ', error occurred while opening file: ' + fh.error);
    } else {
        try {
            var blob = fh.Read(	);
            logger.verbose('blob size: ' + blob.size + ', type: ' + typeof blob);
            try {
                res = util.stGlobal.Decode(blob, format); // "utf-8" seems to be standard, "auto" does not work for me
                logger.verbose('blob -- type: ' + typeof res + ', size: ' + res.length + "\n" + res);
            } catch(e) { logger.error(path + ', st.decode: ' + e.description); }
            blob.Free();
        } catch(e) { logger.error(path + ', fh.read: ' + e.description); }
    }
    fh.Close();
    return res;
}

// internal method
// reads requested file (incl. ADS streams)
function SaveFile(path, contents, format) {
    // path, use "Y:\\Path\\file.txt" or "Y:\\Path\\file.txt:CustomMetaInfo" for ADS
    // contents: any string, e.g. new Date().getTime().toString()
    // format: "base64", "quoted", "auto"=not supplied, "utf-8", "utf-16", "utf-16-le", "utf-16-be"
    // the only ones which worked reliably in my tests are utf-8 & utf-16, since they're the only ones Blob.CopyFrom() supports
    // use only one of encoding.utf8 or encoding.utf16
    //
    // res = SaveFile("Y:\\Path\\file.txt:CustomMetaInfo", encodeURI(new Date().getTime().toString()), encoding.utf16);
    // res = SaveFile("Y:\\Path\\file.txt:CustomMetaInfo", encodeURI("{\"a\": 1}"), encoding.utf8);
    var res = false, decFormat;

    decFormat = format === TEXT_ENCODING.utf16 ? '' : 'utf8';	// unlike ST.Encode()/Decode(), Blob.CopyFrom() uses 'utf8', not 'utf-8'
    format 	  = format === TEXT_ENCODING.utf16 ? 'utf-16' : 'utf-8';

    var fh = DOpus.FSUtil.OpenFile(path, 'wa'); // wa: wa - create a new file, always. If the file already exists it will be overwritten. (This is the default.)
    if(fh.error !== 0) {
        logger.error(path + ', error occurred while opening file: ' + fh.error); return;
    }
    try {
        logger.verbose('String to write to ' + path + ': ' + contents);
        var blob = DOpus.Create.Blob;
        blob.CopyFrom(contents, decFormat);	// seems to use implicitly utf-16, only available optional param is utf8
        res = util.stGlobal.Decode(blob, format);
        logger.verbose('blob -- type: ' + typeof blob + ', size: ' + blob.size + "\n" + res);
        res = fh.Write(blob);
        logger.verbose('Written bytes: ' + res);
    } catch(e) { logger.error(path + ', fh.write: ' + e.description); }
    blob.Free();
    fh.Close();
    return res;
}

// internal method, not used at the moment
// encodes given string to Base64
function ToBase64(contents) {
    var blob = DOpus.Create.Blob;
    blob.CopyFrom(contents);
    var res = util.stGlobal.Encode(blob, 'base64');
    return res;
}

// internal method, not used at the moment
// decodes given Base64 string
function FromBase64(contents) {
    var blob = DOpus.Create.Blob;
    blob.CopyFrom(contents);
    var res = util.stGlobal.Decode(blob, 'base64');
    return res;
}



// internal method
// initializes a DOpus Map for DOpus.Vars cache
function InitCacheIfNecessary() {
    if (!config.get('cache_enabled') || util.sv.Exists('cache')) {
        logger.verbose("Cache not enabled or already initialized -- enabled: " + config.get('cache_enabled'));
        return;
    }
    if (!util.sv.Exists('cache') || !util.sv.Get('cache') || typeof util.sv.Get('cache') === 'undefined') {
        logger.normal('Initializing cache map');
        OnME_ClearCache();
        logger.normal('InitCacheIfNecessary() -- After init, count: ' + util.sv.Get('cache').count);
    }
    logger.error('InitCacheIfNecessary() -- Current count: ' + util.sv.Get('cache').count);
}

// internal method
// debugging method, not used
function DumpCache(prefix, mode) {
    logger.normal((prefix || '') + ' -- Cache dump @' + new Date().getTime().toString());
	var e, item;
    switch (mode) {
        case 1: // names only
            for (e = new Enumerator(util.sv.Get('cache')), item; !e.atEnd(); e.moveNext(), item = e.item()) {
                logger.normal(prefix + ' - name: ' + item);
            }
            break;
        case 2: // count only
            logger.normal(prefix + ' - Count: ' + util.sv.Get('cache').count);
            break;
        default: // everything
            for (e = new Enumerator(util.sv.Get('cache')), item; !e.atEnd(); e.moveNext(), item = e.item()) {
                logger.normal(prefix + ' - name->value: ' + item + ' -> ' + util.sv.Get('cache')(item));
            }
    }
}

// internal method
// reads ADS data, calls ReadFile()
function ReadMetadataADS(oItem) {
    var msn = config.get('MetaStreamName');
    if (!msn) { logger.error('ReadMetadataADS -- Cannot continue without a stream name: ' + msn); return false; }
    // if (!oItem.modify) { logger.error('ReadMetadataADS -- Expected FSUtil Item, got: ' + oItem + ', type: ' + typeof oItem); return false; }
    if ((typeof oItem !== 'object' || typeof oItem.realpath === 'undefined' || typeof oItem.modify === 'undefined')) { logger.error('ReadMetadataADS -- Expected FSUtil Item, got: ' + oItem + ', type: ' + typeof oItem); return false; }
    var rp = ''+oItem.realpath; // realpath returns a DOpus Path object and it does not work well with Map as an object, we need a simple string

    InitCacheIfNecessary();

    // check if cache is enabled and item is in cache
    var res;

    if (config.get('cache_enabled') && util.sv.Get('cache').exists(rp)) {
        logger.verbose(oItem.name + ' found in cache');
        res = util.sv.Get('cache')(rp);
    } else {
        logger.verbose(oItem.name + ' reading from disk');
        res = ReadFile(rp + ':' + msn, TEXT_ENCODING.utf8); // always string or false ion error
        if (res === false) { return res; }
        if (typeof res === 'string' && config.get('cache_enabled') && !util.sv.Get('cache').exists(rp)) {
            logger.verbose(oItem.name + ' was missing in cache, adding');
            res = JSON.parse(res); res.added_to_cache = new Date().getTime().toString(); res = JSON.stringify(res); // this might be overkill
            util.sv.get('cache').set(rp, res);
        }
    }
    // convert to JS object, do not return {} or anything which passes as object but empty string
    return typeof res === 'string' ? JSON.parse(res) : '';
}

// internal method
// saves ADS data, calls SaveFile()
function SaveMetadataADS(oItem, oJSObject) {
    var msn = config.get('MetaStreamName');
    if (!msn) { logger.error('SaveMetadataADS -- Cannot continue without a stream name: ' + msn); return false; }
    if (!oItem.modify) { logger.error('SaveMetadataADS -- Expected FSUtil Item, got: ' + oItem + ', type: ' + typeof oItem); return false; }
    var rp = ''+oItem.realpath; // realpath returns a DOpus Path object, even if its 'default' is map(oItem.realpath) does not work well as key, we need a simple string

    InitCacheIfNecessary();

    var orig_modify = DateToDOpusFormat(oItem.modify);

    util.cmdGlobal.ClearFiles();
    util.cmdGlobal.AddFile(oItem);
    var res = SaveFile(rp + ':' + msn, JSON.stringify(oJSObject), TEXT_ENCODING.utf8);
    if (config.get('keep_orig_modts')) {
        logger.verbose(rp + ', resetting timestamp to: ' + orig_modify);
        util.cmdGlobal.RunCommand('SetAttr MODIFIED "' + orig_modify + '"');
    }
    // check if cache is enabled, add/update unconditionally
    if (config.get('cache_enabled')) {
        oJSObject.added_to_cache = new Date().getTime().toString();
        logger.verbose(oItem.name + ', added to cache: ' + oJSObject.added_to_cache);
        util.sv.get('cache').set(rp, JSON.stringify(oJSObject));
        logger.normal('SaveMetadataADS - Cache count: ' + util.sv.Get('cache').count);
    }
    return res;
}

// internal method
// deletes ADS data, directly deletes "file:stream"
function DeleteMetadataADS(oItem) {
    var msn = config.get('MetaStreamName');
    if (!msn) { logger.error('DeleteMetadataADS -- Cannot continue without a stream name: ' + msn); return false; }
    if (!oItem.modify) { logger.error('DeleteMetadataADS -- Expected FSUtil Item, got: ' + oItem + ', type: ' + typeof oItem); return false; }
    var rp = ''+oItem.realpath; // realpath returns a DOpus Path object, even if its 'default' is map(oItem.realpath) does not work well as key, we need a simple string

    InitCacheIfNecessary();

    var file_stream = oItem.realpath + ':' + msn;
    var orig_modify = DateToDOpusFormat(oItem.modify);

    util.cmdGlobal.ClearFiles();
    util.cmdGlobal.AddFile(oItem);
    util.cmdGlobal.RunCommand('Delete /quiet /norecycle "' + file_stream + '"');
    if (config.get('keep_orig_modts')) {
        logger.verbose(oItem.realpath + ', resetting timestamp to: ' + orig_modify);
        util.cmdGlobal.RunCommand('SetAttr MODIFIED "' + orig_modify + '"');
    }
    if (config.get('cache_enabled')) {
        util.sv.Get('cache').erase(rp);
    }
}


/**
 * helper method to get the Icon Name for development and OSP version
 *
 * do not touch
 */
function GetIcon(scriptPath, iconName) {
    var oPath = DOpus.FSUtil.Resolve(scriptPath);
    var isOSP = oPath.ext === 'osp';
    logger.normal('Requested icon: ' + iconName + ', is OSP: ' + isOSP + '  --  ' + scriptPath);
    return isOSP
        ? '#MExt:' + iconName
        : oPath.pathpart + "\\icons\\ME_32_" + iconName + ".png";
}

/**
 * Date formatter for "SetAttr META lastmodifieddate..."
 *
 * do not touch
 *
 */
function DateToDOpusFormat(oJSDate) {
    return DOpus.Create.Date(oJSDate).Format("D#yyyy-MM-dd T#HH:mm:ss");
}
function SecondsToHHmm(iSeconds) {
    // return DOpus.Create.Date(oJSDate).Format("T#HH:mm");
    // this might be faster
    if (!iSeconds) { return '' }
    var _min = Math.floor(iSeconds / 60);		if (_min < 10) _min = '0' + _min;
    var _sec = Math.round(iSeconds - 60*_min);	if (_sec < 10) _sec = '0' + _sec;
    return _min + ':' + _sec;
}


// internal method
// from https://attacomsian.com/blog/javascript-check-variable-is-object
function isObject(obj) {
    return Object.prototype.toString.call(obj) === '[object Object]';
}




// sprintf - BEGIN
// https://hexmen.com/blog/2007/03/14/printf-sprintf/
{
    // from https://hexmen.com/js/sprintf.js
    /**
	 * JavaScript printf/sprintf functions.
	 *
	 * This code is unrestricted: you are free to use it however you like.
	 *
	 * The functions should work as expected, performing left or right alignment,
	 * truncating strings, outputting numbers with a required precision etc.
	 *
	 * For complex cases these functions follow the Perl implementations of
	 * (s)printf, allowing arguments to be passed out-of-order, and to set
	 * precision and output-length from other argument
	 *
	 * See http://perldoc.perl.org/functions/sprintf.html for more information.
	 *
	 * Implemented flags:
	 *
	 * - zero or space-padding (default: space)
	 *     sprintf("%4d", 3) ->  "   3"
	 *     sprintf("%04d", 3) -> "0003"
	 *
	 * - left and right-alignment (default: right)
	 *     sprintf("%3s", "a") ->  "  a"
	 *     sprintf("%-3s", "b") -> "b  "
	 *
	 * - out of order arguments (good for templates & message formats)
	 *     sprintf("Estimate: %2$d units total: %1$.2f total", total, quantity)
	 *
	 * - binary, octal and hex prefixes (default: none)
	 *     sprintf("%b", 13) ->    "1101"
	 *     sprintf("%#b", 13) ->   "0b1101"
	 *     sprintf("%#06x", 13) -> "0x000d"
	 *
	 * - positive number prefix (default: none)
	 *     sprintf("%d", 3) -> "3"
	 *     sprintf("%+d", 3) -> "+3"
	 *     sprintf("% d", 3) -> " 3"
	 *
	 * - min/max width (with truncation); e.g. "%9.3s" and "%-9.3s"
	 *     sprintf("%5s", "catfish") ->    "catfish"
	 *     sprintf("%.5s", "catfish") ->   "catfi"
	 *     sprintf("%5.3s", "catfish") ->  "  cat"
	 *     sprintf("%-5.3s", "catfish") -> "cat  "
	 *
	 * - precision (see note below); e.g. "%.2f"
	 *     sprintf("%.3f", 2.1) ->     "2.100"
	 *     sprintf("%.3e", 2.1) ->     "2.100e+0"
	 *     sprintf("%.3g", 2.1) ->     "2.10"
	 *     sprintf("%.3p", 2.1) ->     "2.1"
	 *     sprintf("%.3p", '2.100') -> "2.10"
	 *
	 * Deviations from perl spec:
	 * - %n suppresses an argument
	 * - %p and %P act like %g, but without over-claiming accuracy:
	 *   Compare:
	 *     sprintf("%.3g", "2.1") -> "2.10"
	 *     sprintf("%.3p", "2.1") -> "2.1"
	 *
	 * @version 2011.09.23
	 * @author Ash Searle
	 */
    function sprintf() {
        function pad(str, len, chr, leftJustify) {
            var padding = (str.length >= len) ? '' : Array(1 + len - str.length >>> 0).join(chr);
            return leftJustify ? str + padding : padding + str;

        }

        function justify(value, prefix, leftJustify, minWidth, zeroPad) {
            var diff = minWidth - value.length;
            if (diff > 0) {
                if (leftJustify || !zeroPad) {
                    value = pad(value, minWidth, ' ', leftJustify);
                } else {
                    value = value.slice(0, prefix.length) + pad('', diff, '0', true) + value.slice(prefix.length);
                }
            }
            return value;
        }

        var a = arguments, i = 0, format = a[i++];
        return format.replace(sprintf.regex, function (substring, valueIndex, flags, minWidth, _, precision, type) {
            if (substring == '%%') return '%';

            // parse flags
            var leftJustify = false, positivePrefix = '', zeroPad = false, prefixBaseX = false;
            for (var j = 0; flags && j < flags.length; j++) switch (flags.charAt(j)) {
                case ' ': positivePrefix = ' '; break;
                case '+': positivePrefix = '+'; break;
                case '-': leftJustify = true; break;
                case '0': zeroPad = true; break;
                case '#': prefixBaseX = true; break;
            }

            // parameters may be null, undefined, empty-string or real valued
            // we want to ignore null, undefined and empty-string values

            if (!minWidth) {
                minWidth = 0;
            } else if (minWidth == '*') {
                minWidth = +a[i++];
            } else if (minWidth.charAt(0) == '*') {
                minWidth = +a[minWidth.slice(1, -1)];
            } else {
                minWidth = +minWidth;
            }

            // Note: undocumented perl feature:
            if (minWidth < 0) {
                minWidth = -minWidth;
                leftJustify = true;
            }

            if (!isFinite(minWidth)) {
                throw new Error('sprintf: (minimum-)width must be finite');
            }

            if (precision && precision.charAt(0) == '*') {
                precision = +a[(precision == '*') ? i++ : precision.slice(1, -1)];
                if (precision < 0) {
                    precision = null;
                }
            }

            if (precision == null) {
                precision = 'fFeE'.indexOf(type) > -1 ? 6 : (type == 'd') ? 0 : void (0);
            } else {
                precision = +precision;
            }

            // grab value using valueIndex if required?
            var value = valueIndex ? a[valueIndex.slice(0, -1)] : a[i++];
            var prefix, base;

            switch (type) {
                case 'c': value = String.fromCharCode(+value);
                case 's': {
                // If you'd rather treat nulls as empty-strings, uncomment next line:
                // if (value == null) return '';

                    value = String(value);
                    if (precision != null) {
                        value = value.slice(0, precision);
                    }
                    prefix = '';
                    break;
                }
                case 'b': base = 2; break;
                case 'o': base = 8; break;
                case 'u': base = 10; break;
                case 'x': case 'X': base = 16; break;
                case 'i':
                case 'd': {
                    var number = parseInt(+value);
                    if (isNaN(number)) {
                        return '';
                    }
                    prefix = number < 0 ? '-' : positivePrefix;
                    value = prefix + pad(String(Math.abs(number)), precision, '0', false);
                    break;
                }
                case 'e': case 'E':
                case 'f': case 'F':
                case 'g': case 'G':
                case 'p': case 'P':
                {
                    var number = +value;
                    if (isNaN(number)) {
                        return '';
                    }
                    prefix = number < 0 ? '-' : positivePrefix;
                    var method;
                    if ('p' != type.toLowerCase()) {
                        method = ['toExponential', 'toFixed', 'toPrecision']['efg'.indexOf(type.toLowerCase())];
                    } else {
                    // Count significant-figures, taking special-care of zeroes ('0' vs '0.00' etc.)
                        var sf = String(value).replace(/[eE].*|[^\d]/g, '');
                        sf = (number ? sf.replace(/^0+/, '') : sf).length;
                        precision = precision ? Math.min(precision, sf) : precision;
                        method = (!precision || precision <= sf) ? 'toPrecision' : 'toExponential';
                    }
                    var number_str = Math.abs(number)[method](precision);
                    // number_str = thousandSeparation ? thousand_separate(number_str): number_str;
                    value = prefix + number_str;
                    break;
                }
                case 'n': return '';
                default: return substring;
            }

            if (base) {
                // cast to non-negative integer:
                var number = value >>> 0;
                prefix = prefixBaseX && base != 10 && number && ['0b', '0', '0x'][base >> 3] || '';
                value = prefix + pad(number.toString(base), precision || 0, '0', false);
            }
            var justified = justify(value, prefix, leftJustify, minWidth, zeroPad);
            return ('EFGPX'.indexOf(type) > -1) ? justified.toUpperCase() : justified;
        });
    }
    sprintf.regex = /%%|%(\d+\$)?([-+#0 ]*)(\*\d+\$|\*|\d+)?(\.(\*\d+\$|\*|\d+))?([scboxXuidfegpEGP])/g;

    // /**
    //  * Trival printf implementation, probably only useful during page-load.
    //  * Note: you may as well use "document.write(sprintf(....))" directly
    //  */
    // function printf() {
    //     // delegate the work to sprintf in an IE5 friendly manner:
    //     var i = 0, a = arguments, args = Array(arguments.length);
    //     while (i < args.length) args[i] = 'a[' + (i++) + ']';
    //     document.write(eval('sprintf(' + args + ')'));
    // }
}
// sprintf - END




/*
	// not used at the moment, see sprintf
	// as always StackOverflow to the rescue!
	// from https://stackoverflow.com/a/4673436
	if (!String.prototype.format) {
		String.prototype.format = function () {
			var args = arguments;
			return this.replace(/{(\d+)}/g, function (match, number) {
				return typeof args[number] != 'undefined'
					? args[number]
					: match
					;
			});
		};
	}
*/
