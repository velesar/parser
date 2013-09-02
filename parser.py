from lxml import etree
import sys, os, fnmatch
from docx import *

rootdir = sys.argv[1]

def find_files(dir, pattern):
	for root, dirs, files in os.walk(dir):
		for basename in files:
			if fnmatch.fnmatch(basename, pattern):
				filename = os.path.join(root, basename)
				yield filename




for files in find_files(rootdir, '*.xml'):
	filexml = files

result = etree.parse(filexml)
metatags = {}

for metatag in result.xpath("/html/head/meta"):
	metatags[metatag.get('name')] = metatag.get("content")

headers = []

for header in result.xpath('/html/body/div/h2'):
	headers.append(header.text)
headers.pop(0)

divs = []
for div in result.xpath('/html/body/div'):
	divs.append(div)

document = newdocument()
relationships = relationshiplist()

body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
body.append(heading(metatags['booktitle'], 1))
body.append(heading(metatags['author'], 2))
body.append(paragraph('%s (%s)' %(metatags['year'], metatags['firstpub']))) 

body.append(pagebreak(type='page', orient='portrait'))
body.append(heading(metatags['publisher'], 2))
body.append(heading(metatags['address'], 2))
body.append(pagebreak(type='page', orient='portrait'))

for header in headers:
	body.append(paragraph(header, style='ListBullet'))
body.append(pagebreak(type='page', orient='portrait'))

def make_chapters(n):
	body.append(heading(headers[n], 2))
	for div in divs:
		if (div.get('id')[6]) == str(n+1):
			for abstr in div.xpath('p'):
				body.append(paragraph("%s%s"%(abstr.text ,abstr.tail)))
	body.append(pagebreak(type='page', orient='portrait'))

for j in range(len(headers)):
	make_chapters(j)


title = 'prozess'
subject = 'smth'
creator = ''
keywords = []

coreprops = coreproperties(title=title, subject=subject, creator=creator,
                               keywords=keywords)
appprops = appproperties()
contenttypes = contenttypes()
websettings = websettings()
wordrelationships = wordrelationships(relationships)
if not os.path.exists('output'):
        os.makedirs('output')
    # Save our document
savedocx(document, coreprops, appprops, contenttypes, websettings,wordrelationships, 'output/%s %s %s.docx' %(metatags['booktitle'], metatags['author'], metatags['year']))