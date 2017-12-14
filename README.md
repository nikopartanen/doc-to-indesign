# doc-to-publication

Microsoft Word is without doubt the most widely used program to write reports, research articles and books. This means that in proofreading, layout making and editing phases working with Word documents is inevitable. I use InDesign in my semi-professional freelance jobs, so I focus here to that as well.

There are several problems that one can encounter when importing Word documents to publishig software such as InDesign. I assume the biggest issue is usually that the Word document is not actually containing what it seems to contain, so when we import the character and paragraph styles, there is actually lots of stuff that doesn't belong there.

Lots of things are just online without clear license information, so I just add here links and documentation to the things I find work well.

**Note one:** In most of the cases these are edits the document author most likely should not do, unless they are profoundly familiar with what is being discussed here.

**Note two:** Internet is full of professional descriptions of different workflows, I just try to document here what I'm usually doing. What is typical for the texts I work with is that they contains lots of special characters and are often in languages that are not typographically very well supported.

**Note three:** I usually work with texts where the biggest disaster is losing cursive randomly when importing to InDesign. It is nice if styles come in consistently, but making sure some critical formatting is not lost is most important thing.

## Word Macros

- [Remove unused styles, 3rd listed variant](https://answers.microsoft.com/en-us/office/forum/office_2010-word/word-2010-how-does-one-remove-any-unused-styles-in/d9d879ea-d89f-453d-bc8e-6d3dd6f4e48d?auth=1)
    - Makes mapping styles easier
- [Force all fonts to X](https://word.tips.net/T011069_Finding_Text_Not_Using_a_Particular_Font.html)
    - This is useful to make sure there are no fonts hiding somewhere beyond the local overrides that are visible in the document, although I'm not currently sure if this really fixes it to what extent
    - It doesn't seem to work with really messed up documents, but it seems to bring onto surface some underlying previously copy-pasted fonts and things that are there, but which Word usually doesn't show anywhere
- [Remove track-changes before date, 3rd listed variant](https://www.datanumen.com/blogs/4-quick-ways-view-accept-revisions-date-word-document)
    - Invariably the author sends you new version after you have already started to do something. Or maybe there are changes after you have already proofread most of the document. This way it is possible to find easily which changes are new.
- When the .doc document contains some deeply underlying formatting that just doesn't come up easily, but is still messing everything up, saving it in docx-format at times seems to strip away the worst issues, basically it seems to remove some information Word for some bizarre reason stores about ancient formatting layers or something like this.
   - After this you can replace all italics and bolds with character styles, import it all, select all, remove all overrides, and you should have a nice text

## Testing methods

If Pandoc conversion returns correct formatting, this is at least a good sign.

    pandoc -s document.docx -o test.md

## What are the problems?

- After import the text is full of overrides
- Some styles are lost in locations that are almost impossible to splot (i.e. cursive disappears in some rare contexts)

## How to prepare the document

- Remove unused styles
- Map the styles explicitly into InDesign styles while doing Place
- Use Word's Styles > Show Style Guides to find parts that are suspect
- Explicitly mapping the cursive, bold etc. into character styles in Word and InDesign seems to be the best way to make sure you can delete all local overrides

## Where the problems come from?

- Characters, texts and words are copy-pasted around 
- Actually, I think that's the biggest source of all problems…
