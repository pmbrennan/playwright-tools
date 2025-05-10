# Example Tagged Script

Below is an example of a script using the PlaywrightTools tag format. This is what you would write in LibreOffice Writer before applying formatting.

```
## THE SCENE: A COFFEE SHOP - MORNING

## The morning rush is in full swing. ALICE sits at a table with her laptop. 

## BOB enters, spots her, and approaches.

/b/ Hey, Alice! Mind if I join you?

/a/ [[Distracted]] Oh, hi Bob. Sure, have a seat.

/b/ Thanks. (Sits down) So, how's the writing going?

/a/ Ugh. *Terrible*. I've been staring at this blank page for an hour.

/b/ Writer's block?

/a/ More like writer's \massive concrete wall\. I just can't seem to get past the opening scene.

/b/ What's the story about?

/a/ It's a play about two people who meet in a coffee shop. (Laughs) Very _original_, I know.

/b/ Hey, write what you know, right? (Gestures around) This is perfect research.

/a/ I suppose. But I'm struggling with making their dialogue sound natural.

/b/ Well... maybe you could record us talking? Use it as reference?

/a/ [[Perking up]] That's actually not a bad idea.

/b/ See? I'm helpful sometimes.

## ALICE takes out her phone and sets it on the table.

/a/ Okay, now act natural.

/b/ [[Suddenly stiff]] I am acting completely natural. This is how I always sit. And talk. Naturally.

/a/ (Laughs) This isn't going to work if you're this self-conscious.

/b/ Sorry. Let me try again. (Clears throat) So, Alice, tell me about your day.

## The BARISTA calls out from behind the counter.

/c/ Caramel macchiato for Bob!

/b/ That's me. Hold that thought.

## BOB gets up and walks to the counter.

/a/ [[To herself]] This might actually work...

@@ END OF SCENE
```

After running the "Apply Formatting" macro, the above text would be transformed into properly formatted play format with:

1. Character names centered and capitalized
2. Stage directions properly formatted and italicized
3. Text decoration (bold, italic, underline) properly applied
4. Overline stage directions displayed between character names and lines
5. Inline stage directions properly formatted in italics

## Creating a tags.txt File

If you want to set up your character tags in advance, you can create a `tags.txt` file in the same directory as your script with the following content:

```
ALICE,a
BOB,b
BARISTA,c
```

This will associate:
- `/a/` with "ALICE"
- `/b/` with "BOB"
- `/c/` with "BARISTA"

## Common Workflow

1. Write your script using the tag format shown above
2. Run the "Apply Formatting" macro to see how it looks in standard play format
3. Continue writing with tags.
4. You can add text in tag format even to a previously-formatted script. Just
   hit "Apply Formatting" at any time to format newly-added text.
5. Before sharing or printing, run "Apply Formatting" again
6. Use "Highlight Options" to prepare scripts for individual actors

## Helpful Tips

- Keep your character tags short (1-3 letters) for faster typing
- Use consistent capitalization in your stage directions
- Leave a blank line between different characters' dialogue
- Use "Break Up Long Speeches" if a character has a very long monologue
- Remember to save your document regularly!
