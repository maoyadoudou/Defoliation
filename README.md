# Defoliation
## Introduction
Defoliation is my first project after I worked, it is used for generating a MS-word file with a kind of template. 
Defoliation can copy template from a source MS-word file, insert some specific information in the template, gather 
these templates with different information in one MS-word file. 
## Notice
1. **My English is not good enough, some wrong words spelling, bad grammar sentences or confused semantics may include 
   in my narrative. You are welcome to point out my mistake, if you are glad to do. Thank you!** 
2. Defoliation ~~is still not finished~~ has finished basic function, ~~I researched and wrote this project in my privacy 
   exercise project in past, and nowadays~~ if option is 0, it ~~only~~ could copy word information (texts, tables, 
   pictures, equations) and styles ~~, and I will move these codes to this project recently~~. If option is 1, during 
   copying, it could replace parameter with your specific data. But I still not finished manual, sorry!
3. I only have 1 year work experience in Java development in a start-up company, lots of thing that I didn't consider 
   in this project, you are welcome to point it out, thank you.
## License
I used MIT LICENSE, you can share it or edit it base on your requirements. I create this project base on one 
requirement in my work (just one), so I can't consider all the condition, if you have questions or requirements, you 
can also ask me. 
## Basic Manual
1. Creates a XXX.docx template, in your template, every parameter name need be enclose as ${parameter}, if this 
parameter is a picture file path, and you want to display this picture, it should be ${parameter-p}. You can 
skim the pictureTest.docx file, it is my test template.
2. Input your parameter value in dataMap, in the src/main/java/com/maoyadoudou/GenerateDuplicate.java, key in dataMap 
should be same as your parameter name.
3. Input your template file path in sourceFilePath, you can find sourceFilePath 
in src/main/java/com/maoyadoudou/GenerateDuplicate.java , and targetFilePath represents the target file path.
4. If you only want to copy you source file, just make option is 0. If you want to insert parameter in your template 
during copying, make option is 1.
## P.S.
Defoliation still needs improvement, and I will add new functions and make it easier to use in future, and I hope this 
project can give you some help or hints.
