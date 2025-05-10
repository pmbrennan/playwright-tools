# Playwright Tools for LibreOffice

## Overview

Playwright Tools is a package for LibreOffice intended to simplify the task of formatting stage plays and to provide playwrights with an environment that promotes rapid writing. The package also provides an easy means to highlight individual character lines and stage directions, intended for use when printed scripts are prepared for readings or productions.

## The Problem This Solves

Formatting a stage play is a non-trivial task. Time spent formatting character names, stage directions, and other elements is time that could be better spent on the creative process of writing good dialogue.

Many playwrights spend time on mechanical work: writing each character's entire name every time they have a line, and then manually centering and capitalizing it, before typing the dialogue. This breaks the flow of writing and impacts the quality of the work.

## How to Use This Package

As a playwright, you can:

1. Write your script using simple tags instead of fully formatted names and directions
2. Use the macros to instantly format your script for proper presentation
3. Convert back to tag format whenever you want to continue writing

### Example:

Instead of writing (and manually formatting):

```
                    JOSEPH
    This is my line. I am saying my line now.
```

You can simply write:

```
/j/ This is my line. I am saying my line now.
```

Then press the "Apply Formatting" button when you want to transform your text.

## Installation

1. Create a directory for your new play.
2. Configure LibreOffice to enable Python script execution:
   - Go to **Tools > Options > LibreOffice > Security**
   - Click **Macro Security**
   - Set the security level to "Medium" or "Low"
   - Under **Trusted Sources**, add your script directory
3. Copy the file `PlaywrightToolsTemplate.odt` to your directory and
   rename it with the name of your new play. You will use this
   as the basis for your own work. The necessary styles, macros and toolbar 
   buttons have already been installed in the document.

## Formatting Tags

The script supports several types of formatting tags:

1. **CHARACTER LINES**: Surround a character tag with slashes
   - Example: `/a/ Hello, world!` becomes:
   ```
                    AMBROSE
   Hello, world!
   ```

2. **OVERLINE STAGE DIRECTIONS**: Place at the start of a line right after the character tag
   - Example: `/a/ [[Slyly]] Hello, world!` becomes:
   ```
                    AMBROSE
                    (Slyly)
   Hello, world!
   ```

3. **INLINE STAGE DIRECTIONS**: Include within the dialogue
   - Example: `/a/ Hey there! (Winks) Come on over!` becomes:
   ```
                    AMBROSE
   Hey there! (Winks) Come on over!
   ```
   The inline stage directions will be italicized.

4. **STAGE DIRECTIONS BLOCK**: Begin a paragraph with `##`
   - Example: `## AT RISE: The stage is dark.` becomes:
   ```
                            AT RISE: The stage is dark.
   ```

5. **TEXT DECORATION**: Use special characters to mark text formatting
   - `*bold*` for bold text
   - `_underline_` for underlined text
   - `\italic\` for italic text

6. **CENTERED LINES**: Begin a paragraph with `@@`

## Main Functions

1. **Apply Formatting** - Converts all tagged text to properly formatted stage play format
2. **Strip Formatting** - Converts a formatted script back to the simpler tagged format for easier editing
3. **Break Up Long Speeches** - Intelligently splits long character speeches that span page boundaries, adding "(CONT'D)" slugs where needed
4. **Highlight Options** - Provides a dialog to:
   - Remove all highlighting
   - Highlight all stage directions
   - Highlight one specific character's lines
5. **Edit Tags** - Opens the tags.txt file for direct editing of character tag/slug pairs

## Tag and Slug Association

The first time you use a character tag (e.g., `/j/`), the script will ask you for the corresponding character name (slug) to use (e.g., "JOSEPH"). It saves these associations in a file called `tags.txt` in the same folder as your document.

The `tags.txt` file format is simple:
```
JOSEPH,j
MARY,m
```

This allows you to continue using short tags while writing, with the proper character names appearing in the formatted script.

## Highlighted Scripts for Readings

A common task for play readings is to highlight each actor's lines in their copy of the script. This package automates the process:

1. Click the "Highlight Options" button
2. Choose "Highlight Stage Directions" to highlight all stage directions
3. Print this copy for the director
4. Choose "Highlight One Character's Lines" and select a character
5. Print this copy for that actor
6. Repeat for each character

The highlight color is yellow by default.

## Features and Benefits

- **Speed**: Write faster with fewer keystrokes
- **Flow**: Maintain your creative flow without interruptions for formatting
- **Flexibility**: Import and export text easily between tagged and formatted states
- **Highlighting**: Create actor-specific scripts in seconds
- **Pagination**: Automatically handle speech breaks across pages with proper "(CONT'D)" labeling

## Requirements

- LibreOffice Writer 6.0 or higher
- The following paragraph styles should be defined in your document:
  - CHARACTERNAME
  - LINE
  - STAGE_DIRECTION_BLOCK
  - STAGE_DIRECTION_OVERLINE
  - STAGE_DIRECTION_INLINE
  - CENTERED_BLOCK

## Creating a Template

For easiest use, create a LibreOffice Writer template with the required styles:

1. Create a new document
2. Define the styles mentioned above
3. Save as a template (.ott file)
4. Create new scripts by starting from this template

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

## Acknowledgments

- Original code by Patrick M Brennan (pbrennan@gmail.com)

## Tips for Use

1. **Quick writing**: Just focus on the dialogue, using short character tags
2. **Apply formatting**: Format when you want to review the script's appearance
3. **Strip formatting**: Return to tagged format when you want to continue writing quickly
4. **Break long speeches**: Use this when a character has a very long speech that spans multiple pages
5. **Proper style definitions**: Make sure your document has all the required paragraph styles defined

Enjoy a more productive and creative playwriting experience with Playwright Tools!
