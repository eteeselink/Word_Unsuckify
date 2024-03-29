Word_Unsuckify
==============

My homebrewn collection of MS Word macros that make it suck less.

Written for Word 2007, might work in other versions too.

RelativeImage
-------------

Insert-and-link images, but store the link as a relative pathname.

Word 2007, by default, stores all links to images, texts, and so on, as absolute paths.
This makes working together on Word files that have e.g. images a pain: 
effectively, you'll need to just insert (and not link) the image, 
and then manually replace an image every time it changes.
Creating a "Link" (or choosing "Insert and Link") will insert an absolute path reference that will only work on your computer.

*RelativeImage* fixes this. 

You only need this macro to *create* relative-path image links. 
People collaborating on the document do *not* need this macro 
(unless they want to insert images the right way too)

### Installation

Copy&paste the code to your Normal.dotm, and add a button to it to your quick access bar.

### Usage

* Click the button you just added, and you will be shown a file selection dialog. 
* Select a picture in the same directory as the document, or a subdirectory of it.

The macro will insert this picture using some tricks, 
so that if someone else updates the entire document (by Ctrl+A, F9),
the images will be updated correctly, 
even on a different computer.
Note that this will only work if on the other computer, 
the same image files are present,
in the same relative location.


### Background

Word does support, 
through some nasty field magic, 
relative paths. 
The trick is to insert a field like

    { INCLUDEPICTURE "{ FILENAME \p } \\..\\relativeFilename.png" \* MERGEFORMAT }

Here, the nested `FILENAME` field gets the full absolute path of the document, 
including the document file name
(e.g. `C:\Users\George\Documents\Party\Invitation.docx`).
The subsequent `\\..\\` is used to go to the document's path
(e.g. `C:\Users\George\Documents\Party`),
so whatever comes after is a path relative to the document's location.

Inserting this field by hand is a pain, for various reasons.
For example, the nested field (`{ FILENAME \p }`) cannot be created simply by typing `{` and `}`,
because that would be too easy. 
Instead, you have to tell Word that it's a real field by hitting Ctrl+F9 ("add field") 
while editing the outer field.

This macro does the nasty work for you.

To understand what the macro did, 
add an image using the macro,
select the image,
and hit Shift+F9 ("show field codes").
You'll see some codes very much like the `INCLUDEPICTURE` code shown above.
You can edit this by hand, if you want to (if you're a nerd like me).
Hit Shift+F9 again to show the picture again.
