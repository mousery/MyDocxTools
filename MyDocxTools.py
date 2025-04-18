from docx.oxml.ns import qn
from docx.oxml.text.run import CT_R
from docx.oxml.text.font import CT_Fonts
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.text.font import Font
from docx.document import Document as Document_Type

from bisect import bisect_left
from itertools import accumulate

from typing import NewType, Iterable

import re


Language = NewType('Language', str)
CHINESE = Language("c")
ENGLISH = Language("e")
ARABIC  = Language("a")
OTHERS  = Language("o")

def set_run_text(run_element: CT_R | Run, text: str) -> CT_R:
    """
    set the text of run_element to text

    Args:
        run_element (CT_R | Run): the run
        text (str): the text that u want the run element's text to change to

    Returns:
        run_element (CT_R): the same run_element but text changed
    """    
    run_element.text = text
    return run_element

def add_run_after_run(added_run: CT_R | Run | str, after_this_run: CT_R | Run) -> CT_R:
    """
    add the run "added_run" after the run "after_this_run".
    
    if "added_run" is a string, 
    the run added after "after_this_run" will be a run that 
    has the same font as "after_this_run" and text from the string "added_run".

    Args:
        added_run (CT_R | Run | str): added run or added text/string
        after_this_run (CT_R | Run): the added_run will be added next to this run
    
    Returns:
        added_run (CT_R): added run
    """
    if isinstance(added_run, str):
        added_run = set_run_text(after_this_run.__deepcopy__(""), added_run)
    elif isinstance(added_run, Run):
        added_run = added_run._element
        
    if isinstance(after_this_run, Run):
        after_this_run = after_this_run._element
    
    after_this_run.addnext(added_run.__deepcopy__(""))
    
    return added_run

def remove_run(run: CT_R | Run):
    """
    remove the run from its parent paragraph.

    Args:
        run (CT_R | Run): run to be removed
    """    

    if isinstance(run, Run):
        run = run._element
    run.getparent().remove(run)

def split_run_at_string_index(run: CT_R | Run, index: int):
    """
    split the "split_run" at a given index of the run's text, 
    such that the text of first  splited part of the run is original text[:index],
    and       the text of second splited part of the run is original text[index:].

    Args:
        split_run (CT_R | Run): the run to be split
        index (int): the index of the splited position in the text of "split_run"

    Raises:
        IndexError: index should be >0 and <len(split_run.text), otherwise raise IndexError.
    """    

    if isinstance(run, Run):
        run = run._element
    
    text = run.text
    if index <= 0 or index >= len(text): raise IndexError(f"can't cut run \"{text}\", which has length {len(text)}, at index {index}.")

    set_run_text(run, text[:index])
    add_run_after_run(text[index:], run)

def delete_paragraph(paragraph: CT_P | Paragraph):
    """
    remove paragraph from its parent.

    Args:
        paragraph (Paragraph | CT_P): paragraph to be removed
    """    

    if isinstance(paragraph, Paragraph):
        p = paragraph._element
        p.getparent().remove(p)
        paragraph._p = paragraph._element = None
    else:
        p = paragraph
        p.getparent().remove(p)
    
def get_font_name(object: CT_Fonts | CT_R | Font | Run, language: Language) -> str | None:
    """
    get font name from rFont.
    
    In a rFont object,
    it can obtain 4 kinds of fonts, each for 1 of the following 4 languages:
    english, chinese, arabics, other languages.
    
    to get english font name, input "e" (or ENGLISH if you import * from DocxTools) in language arg.
    likewise,
    for chinese font name, input "c" (or CHINESE);
    for arabic  font name, input "a" (or ARABIC);
    for others  font name, input "o" (or OTHERS).

    Args:
        rFont (CT_Fonts | CT_R | Font | Run): the object that u want font name from.
        language (str): the language that you want the font name from.

    Raises:
        ValueError: wrong language arg.

    Returns:
        str: the desired font name.
    """
    
    if isinstance(object, Font):
        rFont = object._element.rPr.rFonts
    elif isinstance(object, Run):
        object = object._element
        rFont = object.rPr.rFonts
    elif isinstance(object, CT_R):
        rFont = object.rPr.rFonts
    
    if rFont is None: return None
    
    if language == CHINESE:
        return rFont.get(qn('w:eastAsia'))
    elif language == ENGLISH:
        return rFont.get(qn('w:ascii'))
    elif language == ARABIC:
        return rFont.get(qn('w:cs'))
    elif language == OTHERS:
        return rFont.get(qn('w:hAnsi'))
    else:
        raise ValueError("language must be c, e, a, or o.\n(stands for chinese, english, arabic or others respectively)")

def set_font_name(object: CT_Fonts | CT_R | Font | Run, new_font_name: str | dict[str], language: Language = ENGLISH) -> CT_Fonts:
    """
    set font name of object.
    
    In a rFont object,
    it can obtain 4 kinds of fonts, each for 1 of the following 4 languages:
    english, chinese, arabics, other languages.
    
    to set english font name, input "e" (or ENGLISH if you import * from DocxTools) in language arg.
    likewise,
    for chinese font name, input "c" (or CHINESE);
    for arabic  font name, input "a" (or ARABIC);
    for others  font name, input "o" (or OTHERS).

    Args:
        object (CT_Fonts | CT_R | Font | Run): the rFont object whose font name u want to set.
        new_font_name (str | dict[str]): if str, set font name to new_font_name;
                                         if dict[name], set font name to new_font_name[old_font_name].
        language (str): the language whose font name u want to set.

    Raises:
        ValueError: wrong language arg.

    Returns:
        CT_Fonts: the rFont object after its font name set.
    """  
    if isinstance(new_font_name, dict):
        replace_font_dict = new_font_name
        old_font_name = get_font_name(object, language)
        if old_font_name is None: return None
        if old_font_name in replace_font_dict.keys():
            new_font_name = replace_font_dict[old_font_name]
            return set_font_name(object, new_font_name, language)
        return None
    elif isinstance(new_font_name, str):
        if isinstance(object, Font):
            rFont = object._element.rPr.get_or_add_rFonts()
        elif isinstance(object, CT_R):
            rFont = object.rPr.get_or_add_rFonts()
        elif isinstance(object, Run):
            rFont = object._element.rPr.get_or_add_rFonts()
            
        if language == CHINESE:
            return rFont.set(qn('w:eastAsia'), new_font_name)
        elif language == ENGLISH:
            return rFont.set(qn('w:ascii'), new_font_name)
        elif language == ARABIC:
            return rFont.set(qn('w:cs'), new_font_name)
        elif language == OTHERS:
            return rFont.set(qn('w:hAnsi'), new_font_name)
        else:
            raise ValueError("language must be c, e, a, or o.\n(stands for chinese, english, arabic or others respectively)")

def isolate_para_runs_by_span(paragraph: Paragraph | CT_P, 
                           span: Iterable[int]) -> tuple[int, int]:
    """
    split some of the given paragraph's runs, such that there exist a slice of paragraph.runs
    whose text contain and only contain the text of paragraph.text[span[0]:span[1]].
    
    return the tuple(starting index of the slice, ending index of the slice + 1).
    
    For example:
    
    I have a paragraph like this:
    string index                                                  0 1 2 3 4 5 6 7 8 9 10 
    paragraph text (separated into runs, indicated by |)          H e l|l o   W|o r l d
    run index                                                       0      1       2
    and the span is [4, 8], which mean I want to isolate paragraph text[4:8] i.e. "o Wo" in separate runs,
    
    then after this function is called, the given paragraph becomes like this:
    string index                                                  0 1 2 3 4 5 6 7 8 9 10
    paragraph text (separated into runs, indicated by |)          H e l|l|o   W|o|r l d
    run index                                                       0   1   2   3   4
    and (2, 4) will be returned, which indicates paragraph.runs[2:4] contains the isolated text "o Wo".

    Args:
        para_element (Paragraph | CT_P): paragraph to be split
        span (Iterable[int]): span of the para_element.text that is wanted in the section, i.e. you are trimming para_element.text[span[0]:span[1]].

    Returns:
        slice of paragraph.runs (tuple[int, int]): (starting index of the section, ending index of the section + 1)
    """    
    
    if isinstance(paragraph, Paragraph):
        paragraph = paragraph._element
    
    para_runs_start_index = [len(run_element.text) for run_element in paragraph.r_lst]
    para_runs_start_index = [0] + list(accumulate(para_runs_start_index))

    trim_start = span[0]                                              # index where the match start
    trim_end   = span[1]                                              # index where the match end
    trim_start_run = bisect_left(para_runs_start_index, trim_start)   # index in list para_runs_start_index if trim_start is to be inserted into the list
    trim_end_run   = bisect_left(para_runs_start_index, trim_end)     # index in list para_runs_start_index if trim_end   is to be inserted into the list

    if trim_end != para_runs_start_index[trim_end_run]:
        split_run_at_string_index(paragraph.r_lst[trim_end_run - 1],
                                            trim_end - para_runs_start_index[trim_end_run - 1])
    
    if trim_start != para_runs_start_index[trim_start_run]:
        split_run_at_string_index(paragraph.r_lst[trim_start_run - 1],
                                            trim_start - para_runs_start_index[trim_start_run - 1])
        trim_end_run += 1
    
    return trim_start_run, trim_end_run

def find(paragraph: CT_P, find: str) -> list[tuple[int, int]]:
    """
    Find text in paragraph.
    
    First it finds the given texts in paragraph with regular expression.
    Then will isolate the found texts in the paragraph such that there exists runs that contain and only contain the found text (for details view MyDocxTools.isolate_para_runs_by_span).
    Finally return list of spans of the runs that contains the found text. length of the list equals to the number of occurance of the found text in the paragraph.

    Args:
        paragraph (CT_P): the paragraph to find text from
        find (str): the found text

    Returns:
        list[tuple[int, int]]: _description_
    """    
    para_text = paragraph.text
    matches = list(re.finditer(find, para_text))
    if matches == []: return []

    spans = [k.span() for k in matches]                  # e.g. if pattern = "a", text = "abc", then span = (0, 1)
    
    for i, span in enumerate(spans):
        spans[i] = isolate_para_runs_by_span(paragraph, span)
        
    return spans

def find_and_replace(paragraph_or_body: CT_P, finds: Iterable[str] | str, replaces: Iterable[str] | str, replaced_font: None) -> None:
    """
    You know... find and replace, from microsoft document.
    Just trying to replicate it with python-docx.
    
    Regular expression compatible.
    
    The font of the replaced text is inherited from the first run in the found text.

    Args:
        paragraph_or_body (CT_P | Paragraph | Document): the paragraph/document to undergo find and replace.
        finds (Iterable[str] | str):      find the strings from this list in the paragraph/document ...
        replaces (Iterable[str] | str):   and replace with the corresponding strings in this list.
    """
    if isinstance(paragraph_or_body, Document_Type):
        document = paragraph_or_body
        for paragraph in document._element.iterfind(".//" + qn("w:p")):
            find_and_replace(paragraph, finds, replaces)
    elif isinstance(paragraph_or_body, Paragraph):
        paragraph = paragraph_or_body
        find_and_replace(paragraph._element, finds, replaces)
    elif isinstance(paragraph_or_body, CT_P):
        paragraph = paragraph_or_body\
        
        if isinstance(finds, str):
            finds = [finds]
        if isinstance(replaces, str):
            replaces = [replaces]
        
        for find_text, replace_text in list(zip(finds, replaces)):

            spans = find(paragraph, find_text)
            
            for trim_start_run, trim_end_run in spans[::-1]:
                set_run_text(paragraph.r_lst[trim_start_run], replace_text)
                for run in paragraph.r_lst[trim_start_run + 1: trim_end_run][::-1]:
                    remove_run(run)
                    

if __name__ == "__main__":
    from docx import Document
    
    document = Document(r"../test.docx")
    print()