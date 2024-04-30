OPTIONS PS=MAX FORMCHAR="|----|+|---+=|-/\<>*" MLOGIC MPRINT SYMBOLGEN noxwait noxsync;

%LET JOB = P02;

LIBNAME IN01 "C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\raw_data\";
LIBNAME IN02 "C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\temp_data\";
LIBNAME temp "C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\temp_data\datasets by SA - for checking\";



data sa_df;
	set IN02.sa_prepare_datasets_by_sa(keep=plancode servicearea mconame
											avg_overall_round 
											avg_experiencehp_round HPRat_rat SAper10kmm_rat
                                			avg_care_round Atc_rat HWDC_rat PDRat_rat
                                			avg_prevention_round PPCpre_rat PPCpost_rat AAP_rat CCS_rat
                                			avg_chronic_round BHcomp_rat CDCcomp_rat);
	servicearea = compress(servicearea, " ");
run;

data official_contact_sa;
	set IN02.contact_mco_info_plancode;
	where program = 'STAR Adult';
	keep MCOname mco_full Service_Area plancode telephone web_eng web_spa Program servicearea;
	rename mconame = mco_official;
	servicearea = compress(service_area, " ");
run;

proc sort data=sa_df; by plancode; run;
proc sort data=official_contact_sa; by plancode; run;

data sa_df_new;
	merge sa_df (in=a)
		official_contact_sa;
	by plancode;

	if a;
run;

proc sort data=sa_df_new;
	by servicearea mco_official;
run;

proc contents data=sa_df_new out=var_list(keep=name type) noprint;
run;



proc format;
	value $sa_order_f
		'avg_overall_round' = '01'

		'avg_experiencehp_round' = '02'
		'HPRat_rat' = '03'
		'SAper10kmm_rat' = '04'

		'avg_care_round' = '05'
		'AtC_rat' = '06'
		'HWDC_rat' = '07'
		'PDRat_rat' = '08'

		'avg_prevention_round' = '09'
		'PPCpre_rat' = '10'
		'PPCpost_rat' = '11'
		'AAP_rat' = '12'
		'CCS_rat' = '13'

		'avg_chronic_round' = '14'
		'BHcomp_rat' = '15'
		'CDCcomp_rat' = '16'

	;
run;


proc sort data=official_contact_sa; by Service_Area mco_official; run;

proc format;
	value norating_f
		. = "No rating†"
		other = [8.1];
		;
run;

%macro transform_data(input_dataset, output_prefix, orderf, prog);

	proc contents data=&input_dataset out=_varlist(keep=name);
	run;

	data _varlist;
		set _varlist;

		if name not in ('mco_full', 'plancode', 'MCONAME', 'SERVICEAREA', 'Program', 'mco_official', 'Service_Area', 'web_eng', 'web_spa', 'telephone');
	run;

	proc sql;
		select name into :varlist separated by ' ' from _varlist;
	quit;

	proc sql;
		select distinct servicearea into :sa_list separated by ' ' from &input_dataset;
	quit;

	%let num_sa = %sysfunc(countw(&sa_list));

	%do i = 1 %to &num_sa;
		%let current_sa = %scan(&sa_list, &i);

		proc transpose data=&input_dataset out=_sa_temp;
			where servicearea = "&current_sa";
			var &varlist;
			id mco_official;
		run;

		proc sql;
			create table &output_prefix._&current_sa as
				select *
					from _sa_temp
						order by put(_NAME_, &orderf..)
			;
		quit;

		/* Save datasets by SDA for checking */
		data temp.&output_prefix._&current_sa;
			set &output_prefix._&current_sa (drop=_label_);
			format  _NUMERIC_ norating_f.;
		run;


		/* drop two headers so that the array can be read without error */
		data &output_prefix._&current_sa;
			set &output_prefix._&current_sa (drop=_name_ _label_);
			format  _NUMERIC_ norating_f.;
		run;

		PROC SQL;
			CREATE TABLE contact_&current_sa AS 
				SELECT mco_full, 
					telephone, 
					Web_eng, 
					Web_spa
				FROM official_contact_&prog
					WHERE servicearea = "&current_sa";
		QUIT;

		PROC SQL;
			CREATE TABLE mco_name_&current_sa AS 
				SELECT mco_official
					FROM &input_dataset
						WHERE servicearea = "&current_sa";
		QUIT;

		proc transpose data=mco_name_&current_sa out=mco_name_&current_sa;
			var mco_official;
			id mco_official;
		run;

	data mco_name_&current_sa;
		set mco_name_&current_sa (drop=_name_ _label_);
	run;

	%end;
%mend transform_data;

%transform_data(SA_DF_new, STAR_Adult, $SA_order_f, sa);



%macro auto_fill_table(prog, servicearea, servicearea_full_name);

	proc sql noprint;
		select count(*) into :num_vars
			from dictionary.columns
				where libname='WORK' and memname=upcase("STAR_&prog._&servicearea");
	quit;

	/* calculate the limit of DDE, tell DDE where should it go*/
	%let end_col = %eval(3 + &num_vars - 1);
	filename &prog dde "Excel|STAR &prog-&servicearea_full_name!r3c3:r18c&end_col" notab;

	data _null_;
		set STAR_&prog._&servicearea;
		file &prog;
		array all_vars {*} _ALL_;

		do i = 1 to dim(all_vars);
			put all_vars{i} '09'x @;
		end;

		put;
	run;

	filename &prog dde "Excel|STAR &prog-&servicearea_full_name!r23c2:r28c5" notab;

	data _null_;
		set CONTACT_&servicearea;
		file &prog;
		put mco_full '09'x telephone '09'x web_eng '09'x web_spa '09'x
		;
	run;
	
	filename &prog dde "Excel|STAR &prog-&servicearea_full_name!r2c3:r2c8" notab;

	data _null_;
		set MCO_Name_&servicearea;
		file &prog;
		array all_vars {*} _ALL_;

		do i = 1 to dim(all_vars);
			put all_vars{i} '09'x @;
		end;

		put;
	run;

%mend auto_fill_table;

/* create DDE macro for filling */
filename ddeopen DDE 'Excel|system';

* template file;
x '"C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\raw_data\bySDA template\2024 MCO Report Cards - bySDA STAR Adult.xlsx"';

/* STAR Adult*/
%auto_fill_table(Adult, BEXAR, Bexar);
%auto_fill_table(Adult, Dallas, Dallas);
%auto_fill_table(Adult, ElPaso, El Paso);
%auto_fill_table(Adult, Harris, Harris);
%auto_fill_table(Adult, Hidalgo, Hidalgo);
%auto_fill_table(Adult, Jefferson, Jefferson);
%auto_fill_table(Adult, Lubbock, Lubbock);
%auto_fill_table(Adult, MRSACENTRal, MRSA Central);
%auto_fill_table(Adult, MRSANORTHEAST, MRSA Northeast);
%auto_fill_table(Adult, MRSAWEST, MRSA West);
%auto_fill_table(Adult, NUECES, Nueces);
%auto_fill_table(Adult, TARRANT, Tarrant);
%auto_fill_table(Adult, TRAVIS, Travis);




data _null_;
	file ddeopen;
	put '[error(false)]';
	put '[save.as("C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Output\MCO Report Cards - SA - bySDA-Final.xlsx")]';
	put '[file.close(false)]';
run;