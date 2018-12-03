/**
*  This example covers a few features: 
*  - Multiple worksheets with multipe queries
*  - Creating Styles and applying them to columns
*  - Worksheet Title (spanning multiple cells)
*  - Conditional Formating for a range of cells in a worksheet
*  - Hyperlinking a cell
*  - Worksheet Level Print Header and Footer
*  - Adding a formula to a column
*  - Sending finished report to a file.
*/


/* Use command CREATE DIRECTORY <directory name> as '<pick a location on your machine>'
   to create a directory for the file.
   The directory name should be passed to this procedure as a parameter.
*/
SET SCAN OFF;

CREATE OR REPLACE PROCEDURE employeeReport_file(p_directory VARCHAR2 := NULL) AS

   -- Notice the special hyperlink function in col1 of salary select statement links the column to the Hiredate worksheet.
   v_sql_salary        VARCHAR2(200) := 'SELECT ExcelDocTypeUtils.createWorksheetLink(''Hiredate'',last_name),first_name,salary FROM hr.employees ORDER BY last_name,first_name';
   v_sql_contact       VARCHAR2(200) := 'SELECT last_name,first_name,phone_number,email  FROM hr.employees ORDER BY last_name,first_name';
   v_sql_hiredate      VARCHAR2(200) := 'SELECT last_name,first_name,to_char(hire_date,''MM/DD/YYYY'') hire_date FROM hr.employees ORDER BY last_name,first_name';

   excelReport         ExcelDocumentType := ExcelDocumentType();
   documentArray       ExcelDocumentLine := ExcelDocumentLine(); 

   v_worksheet_rec     ExcelDocTypeUtils.T_WORKSHEET_DATA := NULL;
   v_worksheet_array   ExcelDocTypeUtils.WORKSHEET_TABLE  := ExcelDocTypeUtils.WORKSHEET_TABLE();

   v_sheet_title       ExcelDocTypeUtils.T_SHEET_TITLE := NULL;

   -- Objects for Defining Document Styles (Optional)

   v_style_def         ExcelDocTypeUtils.T_STYLE_DEF := NULL;
   v_style_array       ExcelDocTypeUtils.STYLE_LIST  := ExcelDocTypeUtils.STYLE_LIST(); 

   -- Object for Defining Conditional Formating (Optional)

   v_condition_rec         ExcelDocTypeUtils.T_CONDITION      := NULL;
   v_condition_array       ExcelDocTypeUtils.CONDITIONS_TABLE := ExcelDocTypeUtils.CONDITIONS_TABLE();
 
   -- Conditions are applied to a range of cells ... there can be more than grouping of format conditions per worksheet.
   v_conditional_format_rec   ExcelDocTypeUtils.T_CONDITIONAL_FORMATS;
   v_conditional_format_array ExcelDocTypeUtils.CONDITIONAL_FORMATS_TABLE := ExcelDocTypeUtils.CONDITIONAL_FORMATS_TABLE();
   
   -- Worksheet Print Headers and Footers.  A Header or Footer can have multiple components
   v_header_rec         ExcelDocTypeUtils.T_WORKSHEET_HF_DATA := NULL;
   v_header_rec_array   ExcelDocTypeUtils.WORKSHEET_HF_TABLE  := ExcelDocTypeUtils.WORKSHEET_HF_TABLE();
   
   v_footer_rec         ExcelDocTypeUtils.T_WORKSHEET_HF_DATA := NULL;
   v_footer_rec_array   ExcelDocTypeUtils.WORKSHEET_HF_TABLE  := ExcelDocTypeUtils.WORKSHEET_HF_TABLE();
   

   v_file        UTL_FILE.FILE_TYPE;

BEGIN

  -- Define Styles (Optional)
    v_style_def.p_style_id     := 'LastnameStyle';
    v_style_def.p_text_color   := 'Red';

    ExcelDocTypeUtils.addStyleType(v_style_array,v_style_def);

    v_style_def := NULL;
    v_style_def.p_style_id          := 'SheetTitleStyle';
    v_style_def.p_align_horizontal  := 'Center';
    v_style_def.p_bold              := 'Y';
    v_style_def.p_text_color        := 'Green';

    ExcelDocTypeUtils.addStyleType(v_style_array,v_style_def);

    v_style_def := NULL;
    v_style_def.p_style_id     := 'FirstnameStyle';
    v_style_def.p_text_color   := 'Blue';

    ExcelDocTypeUtils.addStyleType(v_style_array,v_style_def);
    
    -- Heading Row Style
    v_style_def := NULL;
    v_style_def.p_style_id          := 'HeadingRowStyle';
    v_style_def.p_text_color        := 'Black';
    v_style_def.p_font              := 'Times New Roman';
    v_style_def.p_ffamily           := 'Roman';
    v_style_def.p_fsize             := '10';
    v_style_def.p_bold              := 'Y';
    v_style_def.p_underline         := 'Single';
    v_style_def.p_align_vertical    := 'Bottom';
    v_style_def.p_rotate_text_deg   := '45';
    
    ExcelDocTypeUtils.addStyleType(v_style_array,v_style_def);

    -- Style that includes custom borders around numbers
    v_style_def := NULL;
    v_style_def.p_style_id         := 'NumberStyle';
    v_style_def.p_number_format    := '$###,###,###.00';
    v_style_def.p_align_horizontal := 'Right';
    v_style_def.p_custom_xml         := '<Borders>'||
                                            '<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="3"/>'||
                                            '<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3"/>'||
                                            '<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="3"/>'||
                                            '<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="3"/>'||
                                       '</Borders>';

    ExcelDocTypeUtils.addStyleType(v_style_array,v_style_def);

   -- Define Sheet Title
   v_sheet_title.title      := 'Employee Salary Report';

   -- Must Less than or Equal to the max number of columns on the worksheet.
   v_sheet_title.cell_span  := '3';
   v_sheet_title.style      := 'SheetTitleStyle';

   v_worksheet_rec.title    := v_sheet_title;


   -- Add conditional formating for Salary Ranges ... color code salary amounts
   -- across three different ranges. 
 
   v_condition_rec.qualifier    := 'Between';
   v_condition_rec.value        := '0,5000';
   v_condition_rec.format_style := 'color:red';

   ExcelDocTypeUtils.addConditionType(v_condition_array,v_condition_rec);

   v_condition_rec.qualifier    := 'Between';
   v_condition_rec.value        := '5001,10000';
   v_condition_rec.format_style := 'color:blue';

   ExcelDocTypeUtils.addConditionType(v_condition_array,v_condition_rec);


   v_condition_rec.qualifier    := 'Between';
   v_condition_rec.value        := '10001,1000000';
   v_condition_rec.format_style := 'color:green';

   ExcelDocTypeUtils.addConditionType(v_condition_array,v_condition_rec);

   -- Format range for Column 3 starting at row 2 and going to row 65000 ... 
   v_conditional_format_rec.range      := 'R2C3:R65000C3'; 
   v_conditional_format_rec.conditions := v_condition_array;

   ExcelDocTypeUtils.addConditionalFormatType(v_conditional_format_array,v_conditional_format_rec);

   v_worksheet_rec.worksheet_cond_formats := v_conditional_format_array;
   
   
   -- Create Header Footer Elements
   v_header_rec.hf_type  := ExcelDocTypeUtils.HF_DATE_TIME;
   v_header_rec.position := ExcelDocTypeUtils.HF_RIGHT;
   
   ExcelDocTypeUtils.addHeaderFooterType(v_header_rec_array,v_header_rec);
   
   v_header_rec := NULL;
   v_header_rec.hf_type  := ExcelDocTypeUtils.HF_TEXT;
   v_header_rec.position := ExcelDocTypeUtils.HF_CENTER;
   v_header_rec.text     := 'Employee Report';
   v_header_rec.fontsize := '12';
   
   ExcelDocTypeUtils.addHeaderFooterType(v_header_rec_array,v_header_rec);
   
   v_footer_rec := NULL;
   v_footer_rec.hf_type  := ExcelDocTypeUtils.HF_FILEPATH;
   v_footer_rec.position := ExcelDocTypeUtils.HF_RIGHT;
   
   ExcelDocTypeUtils.addHeaderFooterType(v_footer_rec_array,v_footer_rec);
   
   v_footer_rec := NULL;
   v_footer_rec.hf_type  := ExcelDocTypeUtils.HF_PAGE_NUMBER_PAGES;
   v_footer_rec.position := ExcelDocTypeUtils.HF_CENTER;

   
   ExcelDocTypeUtils.addHeaderFooterType(v_footer_rec_array,v_footer_rec);
   

-- Salary
   
   v_worksheet_rec.worksheet_header := v_header_rec_array;
   v_worksheet_rec.worksheet_footer := v_footer_rec_array;
   
   -- !!! SETTING THE LIST ITEM DELIMITER TO A ':' INSTEAD OF THE DEFAULT ',' !!!
   v_worksheet_rec.worksheet_list_delimiter := ':';
    
   v_worksheet_rec.query                 := v_sql_salary;
   v_worksheet_rec.worksheet_name        := 'Salaries';
   v_worksheet_rec.col_count             := 3;
   v_worksheet_rec.col_width_list        := '25:20:15';
   v_worksheet_rec.col_header_freeze     := TRUE;
   v_worksheet_rec.worksheet_show_gridlines := FALSE;
   v_worksheet_rec.col_header_repeat     := TRUE;
   v_worksheet_rec.col_header_list       := 'Lastname:Firstname:Salary';
   v_worksheet_rec.col_header_style_list := 'HeadingRowStyle:HeadingRowStyle:HeadingRowStyle';
   v_worksheet_rec.col_datatype_list     := 'String:String:Number';
   v_worksheet_rec.col_style_list        := 'LastnameStyle:FirstnameStyle:NumberStyle';
   
   -- Add a SUM formula to the last column. Valid simply formulas include SUM, AVERAGE, COUNT, MIN, MAX
   v_worksheet_rec.col_formula_list  := '::SUM';

   ExcelDocTypeUtils.addWorksheetType(v_worksheet_array,v_worksheet_rec);

   v_worksheet_rec := NULL;

   -- Contact
   v_worksheet_rec.worksheet_header := v_header_rec_array;
   v_worksheet_rec.worksheet_footer := v_footer_rec_array;  
 
 
   -- !!! THE LISTS HERE USE THE DEFAULT ITEM DELIMITER ',' !!!
   v_worksheet_rec.query           := v_sql_contact;
   v_worksheet_rec.worksheet_name  := 'Contact_Info';
   v_worksheet_rec.worksheet_orientation := ExcelDocTypeUtils.WS_ORIENT_LANDSCAPE;
   v_worksheet_rec.col_count       := 4;
   v_worksheet_rec.col_firstcol_freeze := TRUE;
   v_worksheet_rec.col_width_list  := '25,25,25,25';
   v_worksheet_rec.col_header_repeat     := TRUE;
   v_worksheet_rec.col_header_list := 'Lastname,Firstname,Phone,Email';
   v_worksheet_rec.col_style_list    := 'LastnameStyle,FirstnameStyle,,';

   ExcelDocTypeUtils.addWorksheetType(v_worksheet_array,v_worksheet_rec);
   v_worksheet_rec := NULL;

   -- Hiredate
   
   v_worksheet_rec.worksheet_header := v_header_rec_array;
   v_worksheet_rec.worksheet_footer := v_footer_rec_array;
   v_worksheet_rec.query           := v_sql_hiredate;
   v_worksheet_rec.worksheet_name  := 'Hiredate';
   v_worksheet_rec.col_count       := 3;
   v_worksheet_rec.col_width_list  := '25,20,20';
   v_worksheet_rec.col_header_repeat     := TRUE;
   v_worksheet_rec.col_header_list := 'Lastname,Firstname,Hiredate';
   v_worksheet_rec.col_style_list    := 'LastnameStyle,FirstnameStyle,,';

   ExcelDocTypeUtils.addWorksheetType(v_worksheet_array,v_worksheet_rec);

   excelReport := ExcelDocTypeUtils.createExcelDocument(v_worksheet_array,v_style_array);


  -- Write document to a file
  -- Assuming UTL file setting are setup in your DB Instance.
  --  
   documentArray := excelReport.getDocumentData;

   -- Use command CREATE DIRECTORY <directory name> as '<pick a location on your machine>'
   -- to create a directory for the file.
   -- The directory name should be passed to this procedure as a parameter.

   v_file := UTL_FILE.fopen(p_directory,'employeeReport.xls','W',4000);

   FOR x IN 1 .. documentArray.COUNT LOOP
  
     UTL_FILE.put_line(v_file,documentArray(x));
    
   END LOOP;

   UTL_FILE.fclose(v_file);  

END;
/
