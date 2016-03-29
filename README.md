# python_and_excel_vba
This project used python to loop through about 1,369 excel files and write a macro to each file, and save each file with a new name indicating the macro had been completed.

The excel files were extracted from .ssx files, which is a proprietary file extension of a company that makes CPR training manikins. The .ssx files contained CPR session data in xml format. The .ssx files were compressed to zip files that were then decrompressed so the xml files could be opened with excel. Once the excel files had been isolated in their own directory, the ssx_mass_macro_walker was used to parse the xml data into a new worksheet with a certain format specified by the CPR project research directory.

The 5 files in the mass_macro_walker_ssx_demo_files folder constitute a sample of the files that were parsed. Running the ssx_mass_macro_walker.py script on those files will result in the desired output generated in this project.
