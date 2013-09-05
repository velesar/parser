parser
======


## Function

./pyBookConvert <base_folder> <word_template.docx> [output_dir] -m [max_count] –d [folder_depth]

<base_folder>		Folder to scan for .XML files
<word_template.docx>	Word Template for output
[output_dir]	Optional, default is „output“ and files are named „title – year – author.docx“	
-m max_count	Optional, default = 0, mening process unlimited files
-d folder_depth	Optionall, folder depth to scan for XML files (default is 2 subfolders)


###  Scan files

Scan all XML files in <base_folder> and create a list/array with the files names, then process them


###  Convert XML to word_document

Convert XML file to word document based on <word_template.docx>. If there is an image in the book directory (e.g. author/book4/(*.png or *.jpg), use that for the first page.


### Output

Place all processed files in [output_dir], default is „output“. Filenames for the word documents are „Title – Year – Author Name.docx“ (First letter is capitalized of each word)