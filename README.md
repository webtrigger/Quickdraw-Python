# Quickdraw-Python
A python port of Quickdraw, which automatically finds raid targets and triggers for NationStates

## Setup

To run Quickdraw-Python from source, you need a Python 3.6+ interpreter, as well as the [Openpyxl](https://pypi.org/project/openpyxl/) package, which is obtainable through Pip: 

`pip install openpyxl`

Regardless of whether or not you're running from source or the compiled exe, Quickdraw-Python accepts a Spyglass sheet as input, which is expected to be placed in the same directory as the Python script. 

In order to generate a Spyglass sheet, you may either use [Snapsheet](https://aavhrf.github.io/webtrigger/snapsheet.html) or the offline [Spyglass](https://github.com/Aptenodyte/Spyglass/releases)

## Usage

Upon program start, the user will be prompted to supply info. The prompts themselves are fairly self-explanatory. The filters provided should be comma seperated, and the following are provided to you as examples:

Embassy filters:  
```
The Black Hawks, Doll Guldur, Frozen Circle, 3 Guys
```

WFE filters:
```
[url=https://www.forum.the-black-hawks.org/, [url=http://forum.theeastpacific.com,  [url=https://www.nationstates.net/page=dispatch/id=485374], [url=https://discord.gg/XWvERyc, [url=https://forum.thenorthpacific.org, [url=https://discord.gg/Tghy5kW, [url=https://www.westpacific.org
```

After this, the script will open a region page in your browser, and ask you to confirm whether or not the region is a good target. This will be in the format of:  
`(target number). Is  (target) (update time) an acceptable target? (y/n) `  

## Output

The results of the raid are saved into a file called `raid_file.txt`, and the format of the triggers is:  
```
1) target url (target update time)
    a) trigger blank template url (trigger length)
```  
It will also automatically generate a `trigger_list.txt` file, which can be used by [KATT](https://github.com/khronion/KATT).  
