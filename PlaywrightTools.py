#!/usr/bin/env python
# -*- coding: utf-8 -*-
# =====================================================================
# Playwright Tools
# [2025-05-09]
#
# by Patrick M Brennan (pbrennan@gmail.com)
# Copyright (C) 2011-2025 Patrick M Brennan
# 
# Python port by [Your Name Here]
# 
# This is a package for LibreOffice, consisting of a document template 
# and a set of macros intended to simplify the task of formatting
# stage plays, and to provide the playwright with an environment
# which promotes rapid writing. The package also provides an easy
# means to highlight individual character lines and stage directions,
# intended for use when printed scripts are prepared for readings or
# productions.
#
# ********************************************************************
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License 
# along with this program (in the file LICENSE.txt).  If not, see
# <http://www.gnu.org/licenses/>.
# *********************************************************************
#
# THE PROBLEM
#
# Formatting a stage play is a nontrivial task. Time spent fussing
# with character slugs and stage directions is time which could be
# more profitably spent thinking about your characters and writing
# good dialogue.
# 
# From talking to playwrights about their work process, I have come
# to believe that many playwrights needlessly spend time doing
# rather pointless mechanical work: writing each character's entire
# name, every time the character has a line, and then manually
# centering and capitalizing it, before actually typing the
# line. This is not only an enormous waste of time, it also breaks
# the flow of the dialogue in the playwright's head, and may
# therefore actually impact the quality of the work. Automating 
# the process by creating macros which type your character slug 
# for you and center it eliminates some of the tedium, but we
# can do better than that.
#
# Good writing is about emotions, and it's hard to get that right
# if you're concerned with the minutiae of stage formatting.
# Fortunately, computers are really good at handling that kind of
# thing, leaving you free to write!
#
# PYTHON PORT
#
# This Python port maintains all the core functionality of the original BASIC
# package while adapting to the LibreOffice Python API. It allows users to
# easily format their plays using the same tag-based system.

import uno
import os
import re
from com.sun.star.awt import FontSlant
from com.sun.star.awt import FontWeight
from com.sun.star.awt import FontUnderline
from com.sun.star.awt import FontStrikeout
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_YES_NO_CANCEL, BUTTONS_OK_CANCEL
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.beans import PropertyValue
from com.sun.star.text import ControlCharacter
from com.sun.star.container import NoSuchElementException

# Global variables
initialized = False
tags = []
slugs = []
stage_dirs_tag = "## "
centered_line_tag = "@@ "
overline_dirs_open_tag = "[["
overline_dirs_close_tag = "]]"
bold_open_tag = "*"
bold_close_tag = "*"
underline_open_tag = "_"
underline_close_tag = "_"
italic_open_tag = "\\"
italic_close_tag = "\\"
strikethrough_open_tag = "-"
strikethrough_close_tag = "-"
n_lines_per_page = 46
default_style_name = "Default"  # Will be determined at runtime

# Button IDs
YES = 2      # BUTTONS_YES
NO = 3       # BUTTONS_NO
CANCEL = 0   # BUTTONS_CANCEL

def split_csv(input_string):
    """
    Split a comma-separated string into a list of strings
    """
    if not input_string:
        return []
    return [item.strip() for item in input_string.split(",")]

def has_delimiter_enclosed_text(in_string, open_delimiter, close_delimiter, strict=False):
    """
    Determine if the given string contains text enclosed by the given delimiters
    """
    if not in_string:
        return False
        
    first_tag_pos = in_string.find(open_delimiter)
    if first_tag_pos < 0:
        return False
        
    second_tag_pos = in_string.find(close_delimiter, first_tag_pos + len(open_delimiter))
    if second_tag_pos < 0:
        return False
        
    if second_tag_pos <= first_tag_pos:
        return False
        
    if strict:
        # TODO: Add strict checking logic here
        pass
        
    return True

def has_overline_stage_direction(in_string):
    """Determine if the string contains an overline stage direction"""
    return has_delimiter_enclosed_text(in_string, overline_dirs_open_tag, overline_dirs_close_tag)

def has_inline_stage_direction(in_string):
    """Determine if the string contains an inline stage direction"""
    return has_delimiter_enclosed_text(in_string, "(", ")")

def has_bold(in_string):
    """Determine if the string contains bold markup"""
    return has_delimiter_enclosed_text(in_string, bold_open_tag, bold_close_tag, True)

def strip_parentheses(in_string):
    """Strip parentheses from a string"""
    out_string = in_string.strip()
    
    if out_string.startswith("(") and out_string.endswith(")"):
        out_string = out_string[1:-1].strip()
    
    return out_string

def get_script_path():
    """Get the path of the current document"""
    doc = XSCRIPTCONTEXT.getDocument()
    url = doc.getURL()
    if not url:
        return None
        
    # Convert URL to system path
    ctx = XSCRIPTCONTEXT.getComponentContext()
    url_transformer = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.util.URLTransformer", ctx)
    url_struct = uno.createUnoStruct("com.sun.star.util.URL")
    url_struct.Complete = url
    url_transformer.parseStrict(url_struct)
    
    file_access = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.ucb.SimpleFileAccess", ctx)
    
    system_path = file_access.getSystemPathFromFileURL(url)
    return os.path.dirname(system_path)

def read_tags_and_slugs():
    """
    Read the character tag/slug table from tags.txt in the document directory
    Returns True if successful, False otherwise
    """
    global tags, slugs
    
    folder = get_script_path()
    if not folder:
        return False
        
    filename = os.path.join(folder, "tags.txt")
    if not os.path.exists(filename):
        return False
        
    tags = []
    slugs = []
    
    try:
        with open(filename, 'r') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                    
                parts = split_csv(line)
                if len(parts) >= 2:
                    slug = parts[0].upper()
                    tag = f"/{parts[1]}/ "
                    
                    tags.append(tag)
                    slugs.append(slug)
    except:
        return False
        
    return True

def write_tags_and_slugs():
    """
    Write the character tag/slug table to tags.txt in the document directory
    Returns True if successful, False otherwise
    """
    global tags, slugs
    
    folder = get_script_path()
    if not folder:
        return False
        
    filename = os.path.join(folder, "tags.txt")
    
    try:
        with open(filename, 'w') as f:
            for i in range(len(tags)):
                tag = tags[i][1:-2]  # Remove slashes and space from "/tag/ "
                slug = slugs[i]
                f.write(f"{slug},{tag}\n")
    except:
        return False
        
    return True

def tag_exists(tag):
    """Check if a tag exists in the tags list"""
    return tag in tags

def slug_exists(slug):
    """Check if a slug exists in the slugs list"""
    return slug in slugs

def message_box(message, title="", message_type=MESSAGEBOX, buttons=BUTTONS_OK):
    """Display a message box with the given message and title"""
    ctx = XSCRIPTCONTEXT.getComponentContext()
    toolkit = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.Toolkit", ctx)
    parent = XSCRIPTCONTEXT.getDocument().CurrentController.Frame.ContainerWindow
    
    mb = toolkit.createMessageBox(parent, message_type, buttons, title, message)
    return mb.execute()

def input_box(message, title="", default=""):
    """Display an input box with the given message, title, and default value"""
    ctx = XSCRIPTCONTEXT.getComponentContext()
    sm = ctx.ServiceManager
    
    # Create the dialog
    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog.PositionX = 100
    dialog.PositionY = 100
    dialog.Width = 250
    dialog.Height = 80
    dialog.Title = title
    
    # Create the label
    label = dialog.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    label.Name = "Label"
    label.PositionX = 5
    label.PositionY = 5
    label.Width = 240
    label.Height = 20
    label.Label = message
    dialog.insertByName("Label", label)
    
    # Create the text field
    text_field = dialog.createInstance("com.sun.star.awt.UnoControlEditModel")
    text_field.Name = "TextField"
    text_field.PositionX = 5
    text_field.PositionY = 25
    text_field.Width = 240
    text_field.Height = 20
    text_field.Text = default
    dialog.insertByName("TextField", text_field)
    
    # Create the OK button
    ok_button = dialog.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_button.Name = "OKButton"
    ok_button.PositionX = 90
    ok_button.PositionY = 50
    ok_button.Width = 60
    ok_button.Height = 20
    ok_button.Label = "OK"
    ok_button.PushButtonType = 1  # PUSH_DEFAULT (OK)
    dialog.insertByName("OKButton", ok_button)
    
    # Create the dialog control
    control_container = sm.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialog", ctx)
    control_container.setModel(dialog)
    
    # Create the dialog
    toolkit = sm.createInstanceWithContext("com.sun.star.awt.ExtToolkit", ctx)
    control_container.createPeer(toolkit, None)
    
    # Execute the dialog
    result = control_container.execute()
    
    # Get the text field value
    text = control_container.getControl("TextField").Text
    
    # Dispose of the dialog
    control_container.dispose()
    
    return text

def get_default_style_name(doc):
    """Get the default paragraph style name for the document"""
    style_families = doc.StyleFamilies
    para_styles = style_families.getByName("ParagraphStyles")
    style_names = para_styles.ElementNames
    
    # Common names for default paragraph style in different language versions
    candidate_names = ["Standard", "Default Style", "Default", "Normal", 
                       "Predeterminado", "Estándar", "Standard", "Par défaut", 
                       "Normale", "Standaard", "Padrão"]
    
    # First try the candidate names
    for name in candidate_names:
        try:
            para_styles.getByName(name)
            return name
        except NoSuchElementException:
            pass
    
    # If we didn't find a match in the common names,
    # let's try to find the style that has no parent (which is usually the default style)
    for name in style_names:
        style = para_styles.getByName(name)
        # Check if this style has no parent or if it's its own parent
        # This is often the case with the default style
        if hasattr(style, "isAutoUpdate") and style.isAutoUpdate:
            return name
        if hasattr(style, "ParentStyle") and style.ParentStyle == "" and hasattr(style, "isUserDefined") and not style.isUserDefined:
            return name
    
    # If all else fails, return "Default"
    return "Default"

def init_tables():
    """Initialize tables and global variables"""
    global initialized, default_style_name
    
    if not initialized:
        read_tags_and_slugs()
        
        doc = XSCRIPTCONTEXT.getDocument()
        default_style_name = get_default_style_name(doc)
        
        initialized = True

def find_unknown_tags():
    """
    Find and interactively add unregistered tags to the tags table
    Returns True if any new tag/slug pairs were added, False otherwise
    """
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    cursor = text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "^\/([a-zA-Z0-9_\-]{1,8})\/ "
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = True
    
    changed = False
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        if not tag_exists(cursor.String):
            choice = message_box(
                f"I found an unknown tag '{cursor.String}'\nWould you like to replace it?\n(Hit Cancel to exit this loop)",
                "Unknown Tag Found!",
                QUERYBOX,
                BUTTONS_YES_NO_CANCEL
            )
            
            if choice == CANCEL:
                break
            
            if choice == YES:
                new_slug = input_box(
                    f"Enter the new slug for the tag '{cursor.String}'\n(or empty string to skip):\n\n* Note: You do not need to CAPITALIZE the slug.",
                    f"Enter new slug for tag {cursor.String}",
                    ""
                )
                
                if new_slug:
                    new_slug = new_slug.upper()
                    new_tag = cursor.String
                    
                    if slug_exists(new_slug):
                        message_box(
                            f"The Slug '{new_slug}' already exists. Skipping...",
                            "Slug Exists",
                            WARNINGBOX
                        )
                    else:
                        tags.append(new_tag)
                        slugs.append(new_slug)
                        changed = True
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)
    
    if changed:
        write_tags_and_slugs()
    
    return changed

def find_unknown_slugs():
    """
    Find and interactively add unregistered slugs to the slugs table
    Returns True if any new tag/slug pairs were added, False otherwise
    """
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    text = doc.Text
    cursor = text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "^[a-zA-Z0-9_\- \#]{1,50}$"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = True
    
    changed = False
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        new_slug = cursor.String.upper()
        
        if not slug_exists(new_slug):
            choice = message_box(
                f"I found an unknown slug '{new_slug}'\nWould you like to replace it?\n(Hit Cancel to exit this loop)",
                "Unknown Slug Found!",
                QUERYBOX,
                BUTTONS_YES_NO_CANCEL
            )
            
            if choice == CANCEL:
                break
            
            if choice == YES:
                new_tag = input_box(
                    f"Enter the new tag for the slug '{cursor.String}'\n(or empty string to skip):\n\n* Note: You do not need to add the leading and trailing /",
                    f"Enter new tag for slug {new_slug}",
                    ""
                )
                
                if new_tag:
                    new_tag = f"/{new_tag}/ "
                    
                    if tag_exists(new_tag):
                        message_box(
                            f"The Tag '{new_tag}' already exists. Skipping...",
                            "Tag Exists",
                            WARNINGBOX
                        )
                    else:
                        tags.append(new_tag)
                        slugs.append(new_slug)
                        changed = True
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)
    
    if changed:
        write_tags_and_slugs()
    
    return changed

def display_tags_table():
    """Display the tags table in a message box"""
    init_tables()
    
    msg = f"Number of Tags: {len(tags)}\n\n"
    
    for i in range(len(tags)):
        msg += f"Tag: {tags[i]} Slug: {slugs[i]}\n"
    
    message_box(msg, "Tags Table", INFOBOX)

def reload():
    """Reset the initialized flag and reload the tables"""
    global initialized
    initialized = False
    init_tables()

def replace_slug_with_tag(target_string, replace_string):
    """
    Find each instance of target_string by itself on a paragraph;
    replace with replace_string and collapse with following paragraph
    """
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "^" + target_string + "$"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = False
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.setString(replace_string)
        
        # Apply a new paragraph style
        cursor.collapseToStart()
        cursor.goRight(len(replace_string), False)
        cursor.goRight(1, True)
        cursor.setString("")
        cursor.ParaStyleName = default_style_name
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)

def collapse_contd_slugs():
    """
    Find each slug ending with " (CONT'D)" and collapse it with previous paragraph
    """
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "^.*\\(CONT'D\\)$"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = False
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        hit = cursor.getString()
        cursor.setString("")
        
        cursor.collapseToStart()
        cursor.goRight(1, True)
        cursor.setString("")
        
        cursor.collapseToStart()
        cursor.goLeft(1, True)
        hit = cursor.getString()
        if ord(hit) == 10:  # newline character
            cursor.setString("")
        
        cursor.collapseToStart()
        cursor.goLeft(1, True)
        hit = cursor.getString()
        if ord(hit) == 10:  # newline character
            cursor.setString(" ")
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)

def replace_tag_with_text_and_styles(target_string, replace_string, target_style, next_para_style):
    """
    Replace each instance of target_string with replace_string,
    applying target_style to it and next_para_style to the paragraph after.
    """
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = target_string
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = False
    search_desc.SearchCaseSensitive = True
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.setString(replace_string)
        cursor.CharUnderline = FontUnderline.NONE
        cursor.CharWeight = FontWeight.NORMAL
        cursor.CharPosture = FontSlant.NONE
        
        # Apply a new paragraph style
        cursor.collapseToStart()
        cursor.goRight(len(replace_string), False)
        
        # Insert a paragraph break
        doc.Text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
        # Apply the next paragraph style if not already STAGE_DIRECTION_OVERLINE
        if cursor.ParaStyleName != "STAGE_DIRECTION_OVERLINE":
            cursor.ParaStyleName = next_para_style
        
        # Move back and apply the target style
        cursor.goLeft(len(replace_string), False)
        cursor.ParaStyleName = target_style
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)

def character_slugs():
    """
    Replace all tags with proper character slugs and format the
    subsequent lines correctly.
    """
    init_tables()
    
    for i in range(len(tags)):
        replace_tag_with_text_and_styles(
            tags[i], slugs[i], "CHARACTERNAME", "LINE")

def character_tags():
    """
    Replace all character slug lines with the corresponding tags and
    collapse them into a one-line representation of the character's line.
    """
    init_tables()
    
    for i in range(len(tags)):
        replace_slug_with_tag(slugs[i], tags[i])

def format_centered_text():
    """Format paragraphs with centered_line_tag as centered blocks"""
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "^" + centered_line_tag
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = False
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.setString("")
        cursor.collapseToStart()
        cursor.ParaStyleName = "CENTERED_BLOCK"
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)

def unformat_centered_text():
    """Convert centered blocks back to tagged format"""
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    enum = doc.Text.createEnumeration()
    
    while enum.hasMoreElements():
        text_element = enum.nextElement()
        
        if text_element.supportsService("com.sun.star.text.Paragraph"):
            if text_element.ParaStyleName == "CENTERED_BLOCK":
                text_element.ParaStyleName = default_style_name
                text_element.String = centered_line_tag + text_element.String

def format_stage_direction_blocks():
    """Format paragraphs with stage_dirs_tag as stage direction blocks"""
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "^" + stage_dirs_tag
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = False
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.setString("")
        cursor.collapseToStart()
        cursor.ParaStyleName = "STAGE_DIRECTION_BLOCK"
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)

def unformat_stage_direction_blocks():
    """Convert stage direction blocks back to tagged format"""
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    enum = doc.Text.createEnumeration()
    
    while enum.hasMoreElements():
        text_element = enum.nextElement()
        
        if text_element.supportsService("com.sun.star.text.Paragraph"):
            if text_element.ParaStyleName == "STAGE_DIRECTION_BLOCK":
                text_element.ParaStyleName = default_style_name
                text_element.String = stage_dirs_tag + text_element.String

def unformat_manual_block_directions():
    """
    Find paragraphs with large left margin and convert them to
    stage direction blocks with tags
    """
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    enum = doc.Text.createEnumeration()
    
    while enum.hasMoreElements():
        text_element = enum.nextElement()
        
        if text_element.supportsService("com.sun.star.text.Paragraph"):
            style_name = text_element.ParaStyleName
            text = text_element.String
            first_20 = text[:20] if len(text) >= 20 else text
            left_margin = text_element.ParaLeftMargin
            
            if style_name != "STAGE_DIRECTION_BLOCK" and first_20 and left_margin >= 5080:
                text_element.String = stage_dirs_tag + text_element.String
                text_element.ParaStyleName = default_style_name

def format_overline_stage_directions():
    """
    Format overline stage directions (with [[ and ]] tags)
    """
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptors
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = overline_dirs_open_tag
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = False
    search_desc.SearchCaseSensitive = False
    
    search_desc2 = doc.createSearchDescriptor()
    search_desc2.SearchString = overline_dirs_close_tag
    search_desc2.SearchBackwards = False
    search_desc2.SearchRegularExpression = False
    search_desc2.SearchCaseSensitive = False
    
    # Find the first open tag
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.gotoEndOfParagraph(True)
        
        if has_overline_stage_direction(cursor.String):
            cursor.collapseToStart()
            cursor.goRight(len(overline_dirs_open_tag), True)
            cursor.setString("(")
            
            cursor.collapseToEnd()
            cursor = doc.findNext(cursor.End, search_desc2)
            cursor.setString(")")
            cursor.collapseToEnd()
            
            doc.Text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
            
            cursor.ParaStyleName = "LINE"
            cursor.gotoEndOfParagraph(True)
            cursor.setString(cursor.String.strip())
            cursor.collapseToStart()
            cursor.goLeft(1, False)
            cursor.ParaStyleName = "STAGE_DIRECTION_OVERLINE"
            cursor.goRight(1, False)
            cursor.gotoEndOfParagraph(False)
        
        # Find the next open tag
        cursor = doc.findNext(cursor.End, search_desc)

def unformat_overline_stage_directions():
    """
    Convert overline stage directions back to tagged format
    """
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptor for the style
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "STAGE_DIRECTION_OVERLINE"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = False
    search_desc.SearchCaseSensitive = False
    search_desc.SearchStyles = True
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.setString(
            overline_dirs_open_tag +
            strip_parentheses(cursor.String) +
            overline_dirs_close_tag + " "
        )
        
        cursor.gotoEndOfParagraph(False)
        cursor.goRight(1, True)
        cursor.setString("")
        cursor.ParaStyleName = default_style_name
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)

def apply_char_style_to_inline_stage_directions(style):
    """Apply character style to inline stage directions (text in parentheses)"""
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    # Create the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "\\([^\\(]*\\)"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = False
    search_desc.SearchStyles = False
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        if cursor.ParaStyleName == "LINE":
            cursor.CharStyleName = style
        
        # Find the next match
        cursor = doc.findNext(cursor.End, search_desc)

def format_inline_stage_directions():
    """Apply the STAGE_DIRECTION_INLINE style to inline stage directions"""
    apply_char_style_to_inline_stage_directions("STAGE_DIRECTION_INLINE")

def unformat_inline_stage_directions():
    """Remove style from inline stage directions"""
    apply_char_style_to_inline_stage_directions(default_style_name)

def apply_text_decoration_from_markup(open_tag, close_tag, posture=None, underline=None, weight=None, strikeout=None):
    """
    Replace markup with appropriate text formatting
    
    Args:
        open_tag: The opening tag (e.g. '*' for bold)
        close_tag: The closing tag (e.g. '*' for bold)
        posture: FontSlant value for italic formatting
        underline: FontUnderline value for underline formatting
        weight: FontWeight value for bold formatting
        strikeout: FontStrikeout value for strikethrough formatting
    """
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    text_format_cursor = doc.Text.createTextCursor()
    
    # Create the search descriptors
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = open_tag
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = False
    search_desc.SearchCaseSensitive = False
    
    search_desc2 = doc.createSearchDescriptor()
    search_desc2.SearchString = close_tag
    search_desc2.SearchBackwards = False
    search_desc2.SearchRegularExpression = False
    search_desc2.SearchCaseSensitive = False
    
    # Find the first open tag
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.gotoEndOfParagraph(True)
        
        if has_delimiter_enclosed_text(cursor.String, open_tag, close_tag, True):
            cursor.collapseToStart()
            
            cursor.goRight(len(open_tag), True)
            cursor.setString("")
            
            cursor.collapseToEnd()
            
            start_range = cursor.Start
            
            cursor = doc.findNext(cursor.Start, search_desc2)
            
            cursor.collapseToStart()
            cursor.goRight(len(close_tag), True)
            
            cursor.setString("")
            cursor.collapseToEnd()
            
            end_range = cursor.End
            
            text_format_cursor.gotoRange(start_range, False)
            text_format_cursor.gotoRange(end_range, True)
            
            if posture is not None:
                text_format_cursor.CharPosture = posture
            
            if underline is not None:
                text_format_cursor.CharUnderline = underline
            
            if weight is not None:
                text_format_cursor.CharWeight = weight
            
            if strikeout is not None:
                text_format_cursor.CharStrikeout = strikeout
        
        # Find the next open tag
        cursor = doc.findNext(cursor.End, search_desc)

def apply_bold():
    """Apply bold formatting to text within *bold* tags"""
    init_tables()
    apply_text_decoration_from_markup(
        bold_open_tag,
        bold_close_tag,
        weight=FontWeight.BOLD
    )

def apply_underline():
    """Apply underline formatting to text within _underline_ tags"""
    init_tables()
    apply_text_decoration_from_markup(
        underline_open_tag,
        underline_close_tag,
        underline=FontUnderline.SINGLE
    )

def apply_italic():
    """Apply italic formatting to text within \italic\ tags"""
    init_tables()
    apply_text_decoration_from_markup(
        italic_open_tag,
        italic_close_tag,
        posture=FontSlant.ITALIC
    )

def apply_strikeout():
    """Apply strikethrough formatting to text within -strikethrough- tags"""
    init_tables()
    # Currently disabled in the original code
    # apply_text_decoration_from_markup(
    #     strikethrough_open_tag,
    #     strikethrough_close_tag,
    #     strikeout=FontStrikeout.SINGLE
    # )

def replace_text_decoration_with_markup(do_bold=False, do_italic=False, do_underline=False, do_strikeout=False):
    """
    Replace text formatting with markup tags
    
    Args:
        do_bold: Whether to replace bold formatting
        do_italic: Whether to replace italic formatting
        do_underline: Whether to replace underline formatting
        do_strikeout: Whether to replace strikethrough formatting
    """
    init_tables()
    
    # Check that exactly one option is selected
    num_options = sum([do_bold, do_italic, do_underline, do_strikeout])
    if num_options != 1:
        message_box("Wrong Number Of Options", "Error", ERRORBOX)
        return
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    view_cursor = doc.getCurrentController().getViewCursor()
    
    # Set up the search descriptor
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = ".*"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = True
    search_desc.SearchCaseSensitive = False
    
    # Create search attributes based on the selected option
    srch_attr = PropertyValue()
    
    if do_bold:
        srch_attr.Name = "CharWeight"
        srch_attr.Value = FontWeight.BOLD
        open_tag = bold_open_tag
        close_tag = bold_close_tag
    elif do_italic:
        srch_attr.Name = "CharPosture"
        srch_attr.Value = FontSlant.ITALIC
        open_tag = italic_open_tag
        close_tag = italic_close_tag
    elif do_underline:
        srch_attr.Name = "CharUnderline"
        srch_attr.Value = FontUnderline.SINGLE
        open_tag = underline_open_tag
        close_tag = underline_close_tag
    elif do_strikeout:
        srch_attr.Name = "CharStrikeout"
        srch_attr.Value = FontStrikeout.SINGLE
        open_tag = strikethrough_open_tag
        close_tag = strikethrough_close_tag
    
    search_desc.setSearchAttributes([srch_attr])
    
    # Find the first match
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        # Get the cursor's page number
        view_cursor.gotoRange(cursor.Start, False)
        page_num = view_cursor.getPage()
        
        # Get the text and replace it with tagged version
        text = cursor.getString()
        cursor.setString(open_tag + text + close_tag)
        
        # Remove the formatting
        if do_bold:
            cursor.CharWeight = FontWeight.NORMAL
        elif do_italic:
            cursor.CharPosture = FontSlant.NONE
        elif do_underline:
            cursor.CharUnderline = FontUnderline.NONE
        elif do_strikeout:
            cursor.CharStrikeout = FontStrikeout.NONE
        
        cursor.collapseToEnd()
        cursor = doc.findNext(cursor.End, search_desc)

def replace_bold_with_markup():
    """Replace bold formatting with *bold* markup tags"""
    replace_text_decoration_with_markup(do_bold=True)

def replace_italic_with_markup():
    """Replace italic formatting with \italic\ markup tags"""
    replace_text_decoration_with_markup(do_italic=True)

def replace_underline_with_markup():
    """Replace underline formatting with _underline_ markup tags"""
    replace_text_decoration_with_markup(do_underline=True)

def replace_strikeout_with_markup():
    """Replace strikethrough formatting with -strikethrough- markup tags"""
    replace_text_decoration_with_markup(do_strikeout=True)

def apply_formatting(*args):
    """
    Apply formatting for known tags, then scan for unknown tags and
    format them as well.
    """
    character_slugs()
    
    if find_unknown_tags():
        character_slugs()
    
    format_centered_text()
    format_overline_stage_directions()
    format_stage_direction_blocks()
    format_inline_stage_directions()
    apply_bold()
    apply_underline()
    apply_italic()
    # apply_strikeout() - Disabled in original

def strip_formatting(*args):
    """
    Remove all relevant style information and apply tags to text.
    """
    doc = XSCRIPTCONTEXT.getDocument()
    cur_selection = doc.currentSelection
    
    collapse_contd_slugs()
    doc.lockControllers()
    
    replace_strikeout_with_markup()
    replace_bold_with_markup()
    replace_italic_with_markup()
    replace_underline_with_markup()
    unformat_inline_stage_directions()
    unformat_stage_direction_blocks()
    unformat_overline_stage_directions()
    unformat_manual_block_directions()
    unformat_centered_text()
    character_tags()
    
    doc.getCurrentController().select(cur_selection)
    doc.unlockControllers()

def edit_tags(*args):
    """Launch an external editor to edit the tags file."""
    folder = get_script_path()
    if not folder:
        message_box("Could not determine document path.", "Error", ERRORBOX)
        return
    
    filename = os.path.join(folder, "tags.txt")
    
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    
    # Create a dispatch helper
    dispatch_helper = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx)
    
    # Create a URL for the file
    url = uno.systemPathToFileUrl(filename)
    
    # Create a property value for the URL
    prop = PropertyValue()
    prop.Name = "URL"
    prop.Value = url
    
    # Dispatch the URL to open the file
    dispatch_helper.executeDispatch(
        desktop.getCurrentFrame().getComponentWindow(),
        ".uno:OpenFromURL",
        "",
        0,
        [prop]
    )

def get_line_num_from_text_cursor(cursor):
    """Get the line number of a text cursor on the current page"""
    view_cursor = XSCRIPTCONTEXT.getDocument().getCurrentController().getViewCursor()
    view_cursor.gotoRange(cursor, False)
    
    view_cursor.collapseToEnd()
    view_cursor.gotoStartOfLine(False)
    
    line_num = 0
    page_num = view_cursor.getPage()
    
    while view_cursor.goUp(1, False) and view_cursor.getPage() == page_num:
        line_num += 1
    
    return line_num

def get_page_num_from_text_cursor(cursor):
    """Get the page number of a text cursor"""
    view_cursor = XSCRIPTCONTEXT.getDocument().getCurrentController().getViewCursor()
    view_cursor.gotoRange(cursor, False)
    
    view_cursor.collapseToEnd()
    view_cursor.gotoStartOfLine(False)
    
    return view_cursor.getPage()

def get_length_of_speech_from_text_cursor(cursor):
    """
    Get the length of a speech (in lines) starting from a character slug
    """
    view_cursor = XSCRIPTCONTEXT.getDocument().getCurrentController().getViewCursor()
    local_cursor = XSCRIPTCONTEXT.getDocument().Text.createTextCursor()
    view_cursor.gotoRange(cursor, False)
    
    view_cursor.gotoStartOfLine(False)
    length_of_speech = 0
    move_successful = view_cursor.goDown(1, False)
    
    while move_successful:
        local_cursor.gotoRange(view_cursor, False)
        p_style = local_cursor.ParaStyleName
        
        if p_style == "LINE" or p_style == "STAGE_DIRECTION_OVERLINE":
            length_of_speech += 1
            move_successful = view_cursor.goDown(1, False)
        else:
            move_successful = False
    
    return length_of_speech

def current_line(doc):
    """Get the current line number on the page"""
    cur_selection = doc.currentSelection
    doc.lockControllers()
    
    try:
        view_cursor = doc.getCurrentController().getViewCursor()
        view_cursor.collapseToEnd()
        view_cursor.gotoStartOfLine(True)
        n_x = len(view_cursor.getString())
        n_y = 0
        
        page = view_cursor.getPage()
        
        while view_cursor.goUp(1, False) and view_cursor.getPage() == page:
            n_y += 1
    finally:
        doc.getCurrentController().select(cur_selection)
        doc.unlockControllers()
    
    return n_y

def break_up_long_speeches(*args):
    """
    Find and break up long character speeches that span page boundaries
    """
    init_tables()
    collapse_contd_slugs()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    slug_cursor = doc.Text.createTextCursor()
    view_cursor = doc.getCurrentController().getViewCursor()
    
    cur_selection = doc.currentSelection
    doc.lockControllers()
    
    # Create the search descriptor for character names
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "CHARACTERNAME"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = False
    search_desc.SearchCaseSensitive = False
    search_desc.SearchStyles = True
    
    lines_left_on_page = n_lines_per_page
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        character_name = cursor.getString()
        text_length = len(character_name)
        
        if text_length > 1:
            character_name = character_name[:-1]
            text_length -= 1
        
        # Remove " (CONT'D)" if present
        while character_name.upper().endswith(" (CONT'D)"):
            character_name = character_name[:-9]
            text_length -= 9
        
        view_cursor.gotoRange(cursor, False)
        slug_cursor.gotoRange(cursor, False)
        slug_page_num = get_page_num_from_text_cursor(cursor)
        slug_line_num = get_line_num_from_text_cursor(cursor)
        
        length_of_speech = get_length_of_speech_from_text_cursor(cursor)
        
        if lines_left_on_page == -1:
            # Guess how many lines are left on the page
            lines_left_on_page = n_lines_per_page - slug_line_num - 1
        
        # Important parameters coded here
        if length_of_speech > 10 and length_of_speech > lines_left_on_page and lines_left_on_page > 5:
            view_cursor.gotoRange(cursor, False)
            view_cursor.collapseToStart(False)
            view_cursor.goDown(lines_left_on_page - 2, False)
            
            # Find the end of the line
            cur_line_num = current_line(doc)
            while cur_line_num == current_line(doc):
                view_cursor.goRight(1, False)
            
            view_cursor.goLeft(1, False)
            
            cursor.gotoRange(view_cursor, False)
            cursor.collapseToEnd()
            doc.Text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
            
            cursor.ParaStyleName = default_style_name
            
            doc.Text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
            
            cursor.ParaStyleName = "CHARACTERNAME"
            cursor.setString(character_name + " (CONT'D)")
            cursor.collapseToEnd()
            
            doc.Text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
            
            cursor.ParaStyleName = "LINE"
            cursor.gotoEndOfParagraph(True)
            cursor.setString(cursor.getString().strip())
        
        lines_left_on_page = n_lines_per_page - slug_line_num - length_of_speech - 2  # Include the slug and a blank
        
        # Handle the case where we have more than 1 page worth of stuff here
        if lines_left_on_page <= 0:
            lines_left_on_page = n_lines_per_page
        
        cursor = doc.findNext(slug_cursor.End, search_desc)
    
    doc.getCurrentController().select(cur_selection)
    doc.unlockControllers()

def unhighlight_everything(*args):
    """Remove all highlighting in the document"""
    init_tables()
    
    doc = XSCRIPTCONTEXT.getDocument()
    enum = doc.Text.createEnumeration()
    
    while enum.hasMoreElements():
        text_element = enum.nextElement()
        
        if text_element.supportsService("com.sun.star.text.Paragraph"):
            text_element.CharBackTransparent = True

def highlight_block_stage_directions(*args):
    """Highlight only the block stage directions"""
    init_tables()
    unhighlight_everything()
    
    doc = XSCRIPTCONTEXT.getDocument()
    cursor = doc.Text.createTextCursor()
    
    cur_selection = doc.currentSelection
    doc.lockControllers()
    
    # Create the search descriptor for stage direction blocks
    search_desc = doc.createSearchDescriptor()
    search_desc.SearchString = "STAGE_DIRECTION_BLOCK"
    search_desc.SearchBackwards = False
    search_desc.SearchRegularExpression = False
    search_desc.SearchCaseSensitive = False
    search_desc.SearchSimilarity = False
    search_desc.SearchStyles = True
    
    # Set the highlight color to yellow
    applied_color = 0xFFFF00  # RGB(255, 255, 0)
    
    cursor = doc.findFirst(search_desc)
    
    while cursor:
        cursor.CharBackColor = applied_color
        cursor = doc.findNext(cursor.End, search_desc)
    
    doc.getCurrentController().select(cur_selection)
    doc.unlockControllers()

def highlight_character_lines(char_name):
    """Highlight all lines of a specific character"""
    init_tables()
    
    highlight_color = 0xFFFF00  # RGB(255, 255, 0) - Yellow
    in_character_line = False
    
    doc = XSCRIPTCONTEXT.getDocument()
    enum = doc.Text.createEnumeration()
    
    # Loop over all paragraphs
    while enum.hasMoreElements():
        text_element = enum.nextElement()
        
        if text_element.supportsService("com.sun.star.text.Paragraph"):
            if text_element.ParaStyleName == "CHARACTERNAME":
                # Get the character name
                slug = text_element.getString().strip().upper()
                
                # Check against the expected character name
                if char_name.upper() in slug:
                    text_element.CharBackColor = highlight_color
                    in_character_line = True
                else:
                    text_element.CharBackTransparent = True
                    in_character_line = False
            
            elif text_element.ParaStyleName in ("LINE", "STAGE_DIRECTION_OVERLINE"):
                if in_character_line:
                    text_element.CharBackColor = highlight_color
                else:
                    text_element.CharBackTransparent = True
            
            elif text_element.ParaStyleName == "STAGE_DIRECTION_BLOCK":
                text_element.CharBackTransparent = True

def create_highlight_options_dialog():
    """Create a dialog for highlight options"""
    ctx = XSCRIPTCONTEXT.getComponentContext()
    sm = ctx.ServiceManager
    
    # Create the dialog
    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog.PositionX = 100
    dialog.PositionY = 100
    dialog.Width = 250
    dialog.Height = 200
    dialog.Title = "Script Highlighting Options"
    
    # Add option buttons
    option_remove = dialog.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    option_remove.Name = "OptionRemoveAllHighlights"
    option_remove.PositionX = 10
    option_remove.PositionY = 10
    option_remove.Width = 230
    option_remove.Height = 16
    option_remove.Label = "Remove All Highlights"
    option_remove.State = 1  # Selected
    dialog.insertByName("OptionRemoveAllHighlights", option_remove)
    
    option_stage = dialog.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    option_stage.Name = "OptionHighlightStageDirections"
    option_stage.PositionX = 10
    option_stage.PositionY = 30
    option_stage.Width = 230
    option_stage.Height = 16
    option_stage.Label = "Highlight Stage Directions"
    dialog.insertByName("OptionHighlightStageDirections", option_stage)
    
    option_char = dialog.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    option_char.Name = "OptionHighlightCharacter"
    option_char.PositionX = 10
    option_char.PositionY = 50
    option_char.Width = 230
    option_char.Height = 16
    option_char.Label = "Highlight One Character's Lines"
    dialog.insertByName("OptionHighlightCharacter", option_char)
    
    # Add character list box
    list_box = dialog.createInstance("com.sun.star.awt.UnoControlListBoxModel")
    list_box.Name = "CharacterList"
    list_box.PositionX = 30
    list_box.PositionY = 70
    list_box.Width = 210
    list_box.Height = 80
    list_box.Dropdown = True
    dialog.insertByName("CharacterList", list_box)
    
    # Add buttons
    ok_button = dialog.createInstance("com.sun.star.awt.UnoControlButtonModel")
    ok_button.Name = "OKButton"
    ok_button.PositionX = 70
    ok_button.PositionY = 160
    ok_button.Width = 80
    ok_button.Height = 20
    ok_button.Label = "OK"
    ok_button.PushButtonType = 1  # PUSH_DEFAULT (OK)
    dialog.insertByName("OKButton", ok_button)
    
    cancel_button = dialog.createInstance("com.sun.star.awt.UnoControlButtonModel")
    cancel_button.Name = "CancelButton"
    cancel_button.PositionX = 160
    cancel_button.PositionY = 160
    cancel_button.Width = 80
    cancel_button.Height = 20
    cancel_button.Label = "Cancel"
    cancel_button.PushButtonType = 2  # PUSH_CANCEL
    dialog.insertByName("CancelButton", cancel_button)
    
    # Create the control container
    control_container = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    control_container.setModel(dialog)
    
    return control_container

def highlight_options(*args):
    """Show dialog for highlight options and apply the chosen highlighting"""
    init_tables()
    
    # Create and set up the dialog
    dialog = create_highlight_options_dialog()
    toolkit = XSCRIPTCONTEXT.getComponentContext().ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.ExtToolkit", XSCRIPTCONTEXT.getComponentContext())
    dialog.createPeer(toolkit, None)
    
    # Get the dialog controls
    option_remove = dialog.getControl("OptionRemoveAllHighlights")
    option_stage = dialog.getControl("OptionHighlightStageDirections")
    option_char = dialog.getControl("OptionHighlightCharacter")
    char_list = dialog.getControl("CharacterList")
    
    # Fill the character list
    for i, slug in enumerate(slugs):
        char_list.addItem(slug, i)
    
    # If there are characters, select the first one
    if slugs:
        char_list.selectItemPos(0, True)
    
    # Execute the dialog
    result = dialog.execute()
    
    if result == 1:  # OK button
        # Determine which option was selected
        if option_remove.getState():
            option_number = 0
        elif option_stage.getState():
            option_number = 1
        elif option_char.getState():
            option_number = 2
            char_to_highlight = char_list.getSelectedItem()
        
        # Apply the highlighting
        if option_number == 0:
            unhighlight_everything()
        elif option_number == 1:
            highlight_block_stage_directions()
        elif option_number == 2:
            highlight_character_lines(char_to_highlight)
    
    # Dispose of the dialog
    dialog.dispose()

# Entry points for the macros
g_exportedScripts = (
    apply_formatting,
    strip_formatting,
    break_up_long_speeches,
    highlight_options,
    highlight_block_stage_directions,
    unhighlight_everything,
    edit_tags,
)
