Word_Unsuckify
==============

My homebrewn collection of MS Word macros that make it suck less.

Written for Word 2007, might work in other versions too.

RelativeImage
-------------

Word 2007, by default, 
does not at all support storing references to files with relative paths. 
This makes working together on Word files that have e.g. images a pain: 
effectively, you'll need to manually replace an image every time it changes, 
since creating a "Link" will insert an absolute path reference that will only work on your computer.

`RelativeImage` fixes this. 
Copy&paste it to your Normal.dotm, and add a button to your quick access bar.
Run it, and you will be given a file selection dialog. 
You must select a picture in the same directory as the document, or a subdirectory of that.

The macro will insert this picture using some tricks, 
so that if someone else updates the entire document (by Ctrl+A, F9),
the images will be updated correctly 
(assuming that person has access to the same images in the same relative location)

Select the image and hit Shift+F9 ("show field codes") to understand what the macro did for you. 
Hit Shift+F9 again to show the picture again.


### Background

Word does support, 
through some nasty field magic, 
relative paths. 
The trick is to insert a field like

    { INCLUDEPICTURE "{ FILENAME \p } \\..\\relativeFilename.png" \* MERGEFORMAT }

Here, the nested `FILENAME` field gets the full absolute path of the document, 
including the document file name
(e.g. `C:\Users\George\Documents\Party\Invitation.docx`)
`\\..\\` is used to go to the document's path, 
(e.g. `C:\Users\George\Documents\Party`)
so whatever comes after is a path relative to the document's location.

Inserting this field by hand is a pain, for various reasons.
For example, the nested field (`{ FILENAME \p }`) cannot be created simply by typing `{` and `}`,
because that would be too easy. 
Instead, you have to tell Word that it's a real field by hitting Ctrl+F9 (add field) 
while editing the outer field.

This macro does the nasty work for you.
