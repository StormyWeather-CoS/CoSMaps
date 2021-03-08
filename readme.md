There are two scripts here. CoSMaps-PSAW searches the subreddit for any post with the word 'map' in the title or any post with the 'map' flair. CoSMaps-GenerateDoc.py generates a .docx file that plays well with the desired CoS wiki format.

I developed this in python 3.8 and ran it on Spyder. If you need help with running the script or python in general, I recommend the python Discord https://pythondiscord.com/

Please note, there is some laziness inherent in this, because I made these to suit my own needs first. If you're going to use these scripts you may want to make a few tweaks.

The expected usage is this:
1) Start with a spreadsheet with your headings at the top:
Post Title |	Post Date |	Resource URL | Post URL |	Useful | Tag1 | Tag2 | Tag3 | Last Updated:

2) Run CoSMaps-PSAW.py
3) A spreadsheet is generated with info in your first 4 colums, plus the cell to the right of Last Updated: will have the date in it.
4) Go through your results line by line and decide if they are useful. If so, put a y in Column E under Useful. Otherwise put n in Column E.
Example: many posts about maps are roughly "Hey does anyone have a map for Yester Hill?". There's nothing wrong with this post, but it's not a useful resource for other DMs.

5) Tag1, Tag2, and Tag3 are read by CoSMaps-GenerateDoc.py. Tag1 is the main Section heading for your wiki page. E.g. Vallaki. Tag2 is the Subsection heading underneath that. E.g. Wizard of Wines. Tag3 is optional and used for small call-outs like Editor's Choice or Fleshing Out Curse of Strahd.

The only tag that is required is Tag1. If you don't include anything in Tag2, that's fine. It reflects something that doesn't have a specific subsection. Here's the example format:

Inputs:

Tag1: Vallaki Tag2: (blank) Tag3: (blank)

Tag1: Vallaki Tag2: (blank) Tag3: (blank)

Tag1: Vallaki Tag2: Wizard of Wines Tag3: (blank)

Tag1: Vallaki Tag2: Wizard of Wines Tag3: Editor's Choice

Outputs:
## Vallaki

(some link) 

(some link)

### Wizard of Wines

(some link)

(some link) Editor's Choice
