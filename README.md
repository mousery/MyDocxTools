# MyDocTools

It has functions that are working with python-docx at xml level (or rather _element level) that I think is useful to manipulate microsoft documents, including:

1. find_and_replace (regular expression compatible)

    You know... find and replace, from microsoft document. Just trying to replicate it with python-docx.

    Paragraphs in Microsoft document are actually split into sections of text-containing objects called runs, and find and replacing text in paragraphs may not be as intuitive as it sounds. So I wrote find_and_replace to make it easier at least for me.

    It can do find and replace with regular expression, without messing up the font format.

2. get_font_name and set_font_name (including chinese, arabics and others font name)

    For some reason python-docx don't support obtaining chinese, arabics and other font name of text-containing object unless going down to the _element level. I wrote get_font_name, change_font and change_font_from_change_dict to make it easier at least for me.

3. set_run_text
4. add_run_after_run
5. remove_run
6. split_run_at_string_index
7. delete_paragraph
8. isolate_para_runs_by_span