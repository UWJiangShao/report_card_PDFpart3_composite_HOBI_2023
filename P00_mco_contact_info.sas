/* Author: Kira Jiang Shao  */ 
/* Update notes: */
/* TODO: Dec 22, 2023: We found that most of the website provided by HHSC contains a space suffix, 
which will cause the website resolved in brower have a "%20" issue. Therefore, a strip will be
utilized to fix this  */

/* Jan 3 2024. the website space issue is not caused by original dataset, but Excel template. A VBA code 
will be exececute to resolve this issue */
/* Jan 4 2024. HHSC has notified that short name will be used as first page MCO name*/
/* Jan 4 2024. Wellpoint Spanish website is now posted and will be updated*/

/* Update Jan 25 2024 */
/* mco website URL change for UnitedHealthCare Community Health Plan */


OPTIONS PS=MAX FORMCHAR="|----|+|---+=|-/\<>*" MLOGIC MPRINT SYMBOLGEN noxwait noxsync;

LIBNAME IN01 "C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\raw_data\";
LIBNAME IN02 "C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\temp_data\";


proc import datafile="C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\raw_data\MCO Contact Information for Report Cards _2024 UPDATED_dec13.xlsx"
    dbms=XLSX
    out=contact_mco_info
    ;
    sheet = "All programs";
run;

proc contents data = contact_mco_info;
    run;


data IN02.contact_mco_info;
    SET contact_mco_info (Keep = VAR3 VAR2 Service_Area Program Telephone_Number Website__English_ Website__Spanish_ rename=(VAR3 = MCOname VAR2 = mco_full Website__English_ =web_eng Website__Spanish_ = web_spa));
	servicearea = compress(Service_Area, " ");
run;

data CONTACT_MCO_INFO;
	set IN02.CONTACT_MCO_INFO;
	if index(Telephone_Number, '(') > 0 or index(Telephone_Number, ')') > 0 or index(Telephone_Number, '-') > 0 then
        telephone = compress(Telephone_Number, '()- ');
    else
        telephone = Telephone_Number;

	mco_short = scan(mconame, 1);
	
	 if program = 'STAR Kids' then
        program_modified = 'STAR Kids';
    else if program = 'STAR PLUS' then
        program_modified = 'STAR PLUS';
	else if program = 'STAR Adult' or program = 'STAR Child' then
        program_modified = 'STAR';
    else
        program_modified = program; 
run;


/* cross over with MCO contact information */
proc import datafile="..\Data\raw_data\plancode.xlsx"
    dbms=XLSX
    out=plancode
    ;
run;

data plancode_valid;
	set plancode (keep=mconame program plancode servicearea status);
	where status = 'A' and program in ('STAR', 'STAR+PLUS', 'STAR Kids');
	mco_short = scan(mconame, 1);
run;

/*  sort by necessary variables before merging */
proc sort data = plancode_valid; 
by servicearea mco_short program; run;
proc sort data = contact_mco_info; 
by service_area mco_short program_modified; run;

proc contents data=contact_mco_info varnum;
proc contents data=plancode_valid varnum;
run;

data contact_mco_info_plancode;
    length mco_short $37 program_modified $15;
    merge contact_mco_info(in=a) 
        plancode_valid (keep = program servicearea plancode mco_short rename=(servicearea = service_area program = program_modified));
    by service_area mco_short program_modified;
    if a;
run;


 /* mco name change based on HHSC email: Amerigroup --> Wellpoint */
/* Update Jan 4 2024, wellpoint spanish website: https://www.wellpoint.com/es */
data contact_mco_info_plancode;
	set contact_mco_info_plancode;
    if MCOname = 'Amerigroup' then do;
        MCOname = 'Wellpoint';
        mco_full = 'Wellpoint';
        telephone_number = '8337312160';
        telephone = '8337312160';
        Web_eng = 'https://www.wellpoint.com/tx/medicaid';
        web_spa = 'https://www.wellpoint.com/es';
    end;
run;


 /* mco website URL change for UnitedHealthCare Community Health Plan */
/* Update Jan 25 2024 */
data contact_mco_info_plancode;
    set contact_mco_info_plancode;
    if MCOname = 'UnitedHealthcare' then do;
        Web_eng = 'https://www.uhc.com/communityplan/texas/plans';
        web_spa = 'https://es.uhc.com/communityplan/texas/plans';
    end;
run;

 /* mco name change for space issue*/
data IN02.contact_mco_info_plancode;
	set contact_mco_info_plancode;
	mco_short_plus = MCOname;
	mco_full_plus = mco_full;

	   if MCOname = 'Aetna Better Health' then do;
        mco_short_plus = 'Aetna Better Health';
        mco_full_plus = 'Aetna Better Health';
    end;

		   if MCOname = 'Blue Cross and Blue Shield' then do;
        mco_short_plus = 'Blue Cross and Blue Shield';
        mco_full_plus = 'Blue Cross and Blue Shield';
    end;

		   if MCOname = 'Community First' then do;
        mco_short_plus = 'Community First';
        mco_full_plus = 'Community First Health';
    end;

			   if MCOname = 'UnitedHealthcare' then do;
        mco_short_plus = 'United';
        mco_full_plus = 'UnitedHealthcare Community Plan';
    end;

/* eliminate the space for website */

    Web_eng = strip(Web_eng);
    Web_spa = strip(Web_spa);

run;





/*data contact_mco_info_name_changed;*/
/*	set IN02.contact_mco_info_plancode;*/
/*run;*/
/**/
/*/* Check what kind of MCO name we have for full & short*/*/
/*proc sql;*/
/*	select distinct mco_full, MCOname*/
/*	from contact_mco_info_name_changed;*/
/*quit;*/