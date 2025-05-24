# WFM-checker

Simble python parser that takes a list of warframe items, mods or arcanes and prints an excel files with the platinum value and total.
This was created for the Warframe Giveaways group for the purpose of processing large quantities of mod donations (or other item donations) with the intent of attributing the correct equivalent ammount of platinum donated in a faster, more efficient manner.

## How it works

This program accepts a list of items via an input file (input.txt), parses said input via regex, queries the wfm api for prices and out them in a excel file, due to some inconsistencies with naming in wfm objects some items or mods might not be parsed correctly, for example the mod "amar's hatred" in wfm is called amars_hatred whereas "summoner's wrath" is called "summoner's_wrath" or "semi-rifle cannonade" being "semi_rifle_cannonade" whereas "semi-shotgun cannonade" is "shotgun_cannonade", due to this please feel free to open issues with relevant mods or items that are not parsed correctly and they will be fixed, also pull requests adding these ecceptions to the parsing are greatly appreciated.

## How to run this

To run this project it's as simple as cloning the repository with

```bash
git clone https://github.com/GodlySchnoz/WFM-checker.git
```

and then installing the required python packages with 
```bash
pip install requests openpyxl
```
or
```bash
pip install -r requirements.txt
```
