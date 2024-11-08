# PDF Combine Utility 
# 
# Introduction 
The Utility developed for convert TLF's files into *.pdf and combine then according to meta-data file. 
![Test Image 3](assets/images/start.jpg)

# Getting Started
 
1.	Download and run pdf_combine.exe
2.	Select target folder with all TLS's files stored. 
3.	Select metadata file or create one using the template.

|TLF|Title3|Title4|Title5|ProgName|Seq|OutputName|Order|
|--|--|--|--|--|--|--|--|
| L |Listing 1| Listing of Study Disposition |All Enrolled Participants| l_101_disp |1 |L_1  |1  |
| L | Listing 2 | Listing of Demographic Characteristics  |All Enrolled Participants  |l_101_demo  |1 |L_2  |1  |
| F |  Figure 7.1.2|Plot of All Laboratory Values Over Time: Hematology  |All Treated Participants  |g_101_lb1  |1 |F_7_1_2  |3  |
|F| Figure 7.2.2| Plot of All Laboratory Values Over Time: Chemistry| All Treated Participants| g_101_lb1|2|F_7_2_2|4|
4.	Review and fix toc.txt files (if any issues). 

# Tips and Known Issues:
- You cannot have any 0x96 (- dash) characters in the metadata file. If you are using dashes in your Titles or in your file names, replace them in the metadata file before loading it into the utility. Speak with your lead about how to replace these best -- some might be better as colons, some as commas, and some simply can be removed.
Ex:
Intent-to-Treat Population -> Intent to Treat Population
Laboratory - Hematology -> Laboratory: Hematology
(TEAE-SI) -> (TEAE SI)

# Contribute
TODO:
1. Add to usage OpenOffice as convert tool.
2. Restructure codebase
3. Add more tests
4. Add more logging
5. Add more examples
