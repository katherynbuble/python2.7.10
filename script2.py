#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import docx
import string

name = "Cranberry"
in_file_ext = ".docx"
out_file_ext = ".csv"
input_name = name + in_file_ext
d = docx.Document(input_name)
full_text = []

for p in d.paragraphs:
    line = p.text
    full_text.append(line)

t = '\n'.join(full_text)
key_words = [ 'Origin:', 'Fruit:', 'Tree:', 'Plant:', 'Bush:', 'Vine:']
for word in key_words:
    if word in t:
        t = t.replace(word, '<b>'+word+'</b>')

lines = t.splitlines()
# omit the first two lines with crop name and comment
split_lines = [line.split('. ', 1) for line in lines[2:len(lines)]]
output_name = name+out_file_ext

with open(output_name, 'w', encoding='utf-8-sig') as f_out:
    # return object for converting data into delimited strings.
    obj_w = csv.writer(f_out) 
    # write the column headers/fields to the writer data file object.
    obj_w.writerow(('id', 'title', 'body', 'langcode', 'field_mlfruitandnut_crop', 'field_mlfruitandnut_cultivar', 
                    'field_mlfruitandnut_fruit', 'field_mlfruitandnut_origin', 'field_mlfruitandnut_tree', 'field_link_the_site')) 
   
    # max = 5
    count = 1
    for line in split_lines:
        id = str(count)
        count = count + 1
        title = line[0]
        body = line[1]
        print(type(body))
        print(body)
        langcode = "" 
        
        field_mlfruitandnut_crop = ""
        field_mlfruitandnut_cultivar = ""
        field_link_the_site = ""
        if body.startswith('See '):
            field_mlfruitandnut_crop = lines[0].title()
            field_mlfruitandnut_cultivar =  title
            field_link_the_site = (body.strip('.')).lstrip('See ') + " (" + title + ")"
        else:
            field_mlfruitandnut_crop = lines[0].title()
            field_mlfruitandnut_cultivar =  title
     
            field_mlfruitandnut_fruit = "" 
        field_mlfruitandnut_origin = ""
        field_mlfruitandnut_tree = "" 
        tline = [id, title, body, langcode, field_mlfruitandnut_crop, field_mlfruitandnut_cultivar, field_mlfruitandnut_fruit, 
                 field_mlfruitandnut_origin, field_mlfruitandnut_tree, field_link_the_site]
        t = () + (tline,)
        obj_w.writerows(t)
