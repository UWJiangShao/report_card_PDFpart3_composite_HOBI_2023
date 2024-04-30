* %include 'C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\2. Admin\Program\P33_export_admin_for_composite.sas';
* %include 'C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\3. Survey\Program\P32_export_survey_for_composite.sas';
* %include 'C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\4. Complaint\Program\P04_export_complaint_for_composite.sas';


OPTIONS PS=MAX FORMCHAR="|----|+|---+=|-/\<>*" MLOGIC MPRINT SYMBOLGEN noxwait noxsync;
%LET JOB = P01;

LIBNAME Admin "..\Data\raw_data\admin\";
LIBNAME Survey "..\Data\raw_data\survey\";
LIBNAME Comp "..\Data\raw_data\complaint\";
LIBNAME temp "..\Data\temp_data\";

/* create four covered program's datasets: */
%macro combine_datasets(prog);

	data &prog._all;
		merge Admin.&prog._admin 
			Survey.&prog._survey 
			Comp.&prog._comp;
		by plancode;
	run;

%mend combine_datasets;

%combine_datasets(sc);
%combine_datasets(sa);
%combine_datasets(sp);
%combine_datasets(sk);

/* ------------------------------------------- import plancode description ------------------------------------------------------------*/
proc import datafile="..\Data\raw_data\plancode.xlsx"
	dbms=XLSX
	out=plancode
;
run;

%macro add_mco_info(prog);
	proc sort data=plancode; by plancode; run;
	proc sort data=&prog._all; by plancode; run;

	data &prog._all;
		merge &prog._all(in=a) 
			plancode (keep=MCONAME PLANCODE SERVICEAREA);
		by plancode;
		if a;
	run;

	proc sort data=&prog._all; by servicearea mconame; run;
%mend add_mco_info;

%add_mco_info(SA);
%add_mco_info(SC);
%add_mco_info(SP);
%add_mco_info(SK);


/* ------------------------------------------- calculate average ratings ------------------------------------------------------------*/
/* This macro will generate two averge, one for showing on each domain, another unrounded is for calculating the final overall score */
%macro calculate_avg(dataset, domain_name, var_list);

	data &dataset;
		set &dataset;
		array vars {*} &var_list;
		%let n_measures = %sysfunc(countw(&var_list));

		if nmiss(of vars{*}) <= (&n_measures / 2) then
			do;
				avg_&domain_name._unround = mean(of vars{*});
				avg_&domain_name._round = round(avg_&domain_name._unround, 0.5);
			end;
		else
			do;
				avg_&domain_name._unround = .;
				avg_&domain_name._round = .;
			end;
	run;

%mend calculate_avg;

%calculate_avg(sc_all, experiencehp,     HPrat_rat SCper10kmm_rat);
%calculate_avg(sc_all, care,      		GCQ_rat HWDC_rat PDrat_rat);
%calculate_avg(sc_all, prevention,      W30comp_rat_unround WCVcomp_rat_unround vacc_rat_unround);
%calculate_avg(sc_all, chronic,        AMR_rat ADD_rat);
%calculate_avg(sc_all, overall,        avg_experiencehp_unround avg_care_unround avg_prevention_unround avg_chronic_unround);

%calculate_avg(sa_all, experiencehp,      HPrat_rat SAper10kmm_rat);
%calculate_avg(sa_all, care,      		ATC_rat HWDC_rat PDrat_rat);
%calculate_avg(sa_all, prevention,      PPCpre_rat PPCpost_rat AAP_rat CCS_rat);
%calculate_avg(sa_all, chronic,         BHcomp_rat_unround CDCcomp_rat_unround);
%calculate_avg(sa_all, overall,        avg_experiencehp_unround avg_care_unround avg_prevention_unround avg_chronic_unround);

%calculate_avg(sp_all, experiencehp,      HPrat_rat SPper10kmm_rat);
%calculate_avg(sp_all, care,      		ATC_rat HWDC_rat PDrat_rat);
%calculate_avg(sp_all, prevention,      AAP_rat cancer_rat_unround);
%calculate_avg(sp_all, chronic,         BHcomp_rat_unround  IET_rat COPD_rat_unround CDCcomp_rat_unround);
%calculate_avg(sp_all, overall,        avg_experiencehp_unround avg_care_unround avg_prevention_unround avg_chronic_unround);

%calculate_avg(sk_all, experiencehp,      HPrat_rat SKper10kmm_rat);
%calculate_avg(sk_all, care,      		Atc_rat WCVcomp_rat_unround SpecTher_rat APM_survey_rat);
%calculate_avg(sk_all, prevention,      coord_rat GNI_rat transit_rat);
%calculate_avg(sk_all, chronic,         BHcoun_rat FUH_rat APM_admin_rat);
%calculate_avg(sk_all, overall,        avg_experiencehp_unround avg_care_unround avg_prevention_unround avg_chronic_unround);










/* export dataset to make bySDA tables */
%macro export_by_sda(prog);
		data temp.&prog._prepare_datasets_by_sa;
			set &prog._all;
		run;
%mend export_by_sda;

%export_by_sda(sc);
%export_by_sda(sa);
%export_by_sda(sp);
%export_by_sda(sk);

** ---- Create frequency table for the rating guide ---------------------------------------------;
%macro create_freq_table (prog);

	proc freq data=&prog._all nlevels;
		table Avg_ExperienceHP_round /list out=RG_&prog._1;
		table Avg_Care_round /list out=RG_&prog._2;
		table Avg_Prevention_round /list out=RG_&prog._3;
		table Avg_Chronic_round /list out=RG_&prog._4;
		table Avg_overall_round /list out=RG_&prog._5;
	run;

%mend create_freq_table;

%create_freq_table(SC);
%create_freq_table(SA);
%create_freq_table(SP);
%create_freq_table(SK);











proc format;
	value norating_f
		. = "--"
		other = [8.1];
	;
run;

** ---- Exporting using DDE --------------------------------------------------------------------;
filename ddeopen DDE 'Excel|system';

* template file;
x '"..\Data\raw_data\composite_template\2024 MCO Report Cards Composites Ratings-unrounded_QL.xlsx"';

/* STAR Child*/
filename SC dde "Excel|STAR Child-ExperienceHP!r3c1:r46c7" notab;

data _null_;
	set sc_all;
	file SC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_rat '09'x SCper10kmm_rat '09'x Avg_ExperienceHP_unround  '09'x Avg_ExperienceHP_round '09'x
	;
	format HPRat_rat SCper10kmm_rat Avg_ExperienceHP_unround Avg_ExperienceHP_round norating_f.;
run;

filename SC dde "Excel|STAR Child-Care!r3c1:r46c8" notab;

data _null_;
	set sc_all;
	file SC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		GCQ_rat '09'x HWDC_rat '09'x PDRat_rat '09'x Avg_Care_unround  '09'x  Avg_Care_round '09'x 
	;
	format GCQ_rat HWDC_rat PDRat_rat  Avg_Care_unround Avg_Care_round norating_f.;
run;

filename SC dde "Excel|STAR Child-Prevention!r3c1:r46c11" notab;

data _null_;
	set sc_all;
	file SC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		W30comp_rat_unround '09'x  W30comp_rat '09'x 
		WCVcomp_rat_unround '09'x  WCVcomp_rat '09'x 
		vacc_rat_unround '09'x  vacc_rat '09'x 
		Avg_Prevention_unround '09'x  Avg_Prevention_round '09'x 
	;
	format  W30comp_rat_unround W30comp_rat
			WCVcomp_rat_unround WCVcomp_rat
			vacc_rat_unround vacc_rat 
			Avg_Prevention_unround Avg_Prevention_round norating_f.;
run;

filename SC dde "Excel|STAR Child-Chronic!r3c1:r46c7" notab;

data _null_;
	set sc_all;
	file SC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		AMR_rat '09'x ADD_rat '09'x 
		Avg_Chronic_unround  '09'x 	Avg_Chronic_round '09'x 
	;
	format AMR_rat ADD_rat 
			Avg_Chronic_unround Avg_Chronic_round norating_f.;
run;

filename SC dde "Excel|STAR Child-Overall!r3c1:r46c13" notab;

data _null_;
	set sc_all;
	file SC;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		Avg_ExperienceHP_unround '09'x 
		Avg_ExperienceHP_round '09'x 

		Avg_Care_unround '09'x 
		Avg_Care_round '09'x 

		Avg_Prevention_unround '09'x 
		Avg_Prevention_round '09'x 

		Avg_Chronic_unround '09'x 
		Avg_Chronic_round '09'x 

		avg_overall_unround '09'x 
		avg_overall_round '09'x 
	;
	format  Avg_ExperienceHP_unround Avg_ExperienceHP_round  
			Avg_Care_unround Avg_Care_round
			Avg_Prevention_unround Avg_Prevention_round 
			Avg_Chronic_unround Avg_Chronic_round 
			avg_overall_unround avg_overall_round norating_f.;
run;




/* ---------------------------------STAR Adult--------------------------------------------------------------------- */
filename SA dde "Excel|STAR Adult-ExperienceHP!r3c1:r46c7" notab;

data _null_;
	set sa_all;
	file SA;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_rat '09'x SAper10kmm_rat '09'x Avg_ExperienceHP_unround  '09'x Avg_ExperienceHP_round '09'x
	;
	format HPRat_rat SAper10kmm_rat Avg_ExperienceHP_unround Avg_ExperienceHP_round norating_f.;
run;

filename SA dde "Excel|STAR Adult-Care!r3c1:r46c8" notab;

data _null_;
	set sa_all;
	file SA;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		AtC_rat '09'x HWDC_rat '09'x PDRat_rat '09'x Avg_Care_unround '09'x  Avg_Care_round '09'x 
	;
	format AtC_rat HWDC_rat PDRat_rat Avg_Care_unround Avg_Care_round norating_f.;
run;

filename SA dde "Excel|STAR Adult-Prevention!r3c1:r46c9" notab;

data _null_;
	set sa_all;
	file SA;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		PPCpre_rat '09'x PPCpost_rat '09'x AAP_rat '09'x CCS_rat '09'x Avg_Prevention_unround '09'x  Avg_Prevention_round '09'x 
	;
	format PPCpre_rat PPCpost_rat AAP_rat CCS_rat Avg_Prevention_unround Avg_Prevention_round norating_f.;
run;

filename SA dde "Excel|STAR Adult-Chronic!r3c1:r46c9" notab;

data _null_;
	set sa_all;
	file SA;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		BHcomp_rat_unround '09'x BHcomp_rat '09'x 
		CDCcomp_rat_unround '09'x CDCcomp_rat '09'x 
		Avg_Chronic_unround '09'x Avg_Chronic_round '09'x 
	;
	format BHcomp_rat_unround BHcomp_rat CDCcomp_rat_unround CDCcomp_rat Avg_Chronic_unround Avg_Chronic_round norating_f.;
run;

filename SA dde "Excel|STAR Adult-Overall!r3c1:r46c13" notab;

data _null_;
	set sa_all;
	file SA;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		Avg_ExperienceHP_unround '09'x 
		Avg_ExperienceHP_round '09'x 

		Avg_Care_unround '09'x 
		Avg_Care_round '09'x 

		Avg_Prevention_unround '09'x 
		Avg_Prevention_round '09'x 

		Avg_Chronic_unround '09'x 
		Avg_Chronic_round '09'x 

		avg_overall_unround '09'x 
		avg_overall_round '09'x 
	;
	format  Avg_ExperienceHP_unround Avg_ExperienceHP_round  
			Avg_Care_unround Avg_Care_round 
			Avg_Prevention_unround Avg_Prevention_round 
			Avg_Chronic_unround Avg_Chronic_round 
			avg_overall_unround avg_overall_round norating_f.;
run;




/* ----------------------------------------STAR+PLUS-------------------------------------------------------- */

filename SP dde "Excel|STAR+PLUS-ExperienceHP!r3c1:r46c7" notab;

data _null_;
	set sp_all;
	file SP;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_rat '09'x SPper10kmm_rat '09'x Avg_ExperienceHP_unround '09'x Avg_ExperienceHP_round '09'x
	;
	format HPRat_rat SPper10kmm_rat Avg_ExperienceHP_unround Avg_ExperienceHP_round norating_f.;
run;

filename SP dde "Excel|STAR+PLUS-Care!r3c1:r46c8" notab;

data _null_;
	set sp_all;
	file SP;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		ATC_rat '09'x HWDC_rat '09'x PDRat_rat '09'x Avg_Care_unround  '09'x  Avg_Care_round '09'x 
	;
	format ATC_rat HWDC_rat PDRat_rat  Avg_Care_unround Avg_Care_round norating_f.;
run;

filename SP dde "Excel|STAR+PLUS-Prevention!r3c1:r46c8" notab;

data _null_;
	set sp_all;
	file SP;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		AAP_rat '09'x cancer_rat_unround  '09'x  cancer_rat '09'x Avg_Prevention_unround  '09'x   Avg_Prevention_round '09'x 
	;
	format AAP_rat cancer_rat  cancer_rat_unround Avg_Prevention_unround Avg_Prevention_round norating_f.;
run;

filename SP dde "Excel|STAR+PLUS-Chronic!r3c1:r46c12" notab;

data _null_;
	set sp_all;
	file SP;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		 BHcomp_rat_unround '09'x BHcomp_rat '09'x
		IET_rat '09'x 
		COPD_rat_unround  '09'x COPD_rat '09'x 
		CDCcomp_rat_unround  '09'x CDCcomp_rat '09'x 
		Avg_Chronic_unround  '09'x Avg_Chronic_round '09'x 
	;
	format  BHcomp_rat_unround BHcomp_rat 
			IET_rat  
			COPD_rat_unround COPD_rat  
			CDCcomp_rat_unround CDCcomp_rat 
			Avg_Chronic_unround Avg_Chronic_round norating_f.;
run;

filename SP dde "Excel|STAR+PLUS-Overall!r3c1:r46c13" notab;

data _null_;
	set sp_all;
	file SP;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		Avg_ExperienceHP_unround '09'x Avg_ExperienceHP_round '09'x 
		Avg_Care_unround '09'x  Avg_Care_round '09'x 
		Avg_Prevention_unround '09'x Avg_Prevention_round '09'x 
		Avg_Chronic_unround '09'x Avg_Chronic_round '09'x 
		avg_overall_unround '09'x avg_overall_round '09'x 
	;
	format Avg_ExperienceHP_round Avg_ExperienceHP_unround 
		   Avg_Care_round Avg_Care_unround 
		   Avg_Prevention_round Avg_Prevention_unround
		   Avg_Chronic_round Avg_Chronic_unround
		   avg_overall_unround avg_overall_round norating_f.;
run;




/*------------------------------------- STAR Kids----------------------------------------------------------- */
filename SK dde "Excel|STAR Kids-ExperienceHP!r3c1:r46c7" notab;

data _null_;
	set sk_all;
	file SK;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		HPRat_rat '09'x SKper10kmm_rat '09'x 
		Avg_ExperienceHP_unround '09'x Avg_ExperienceHP_round '09'x
	;
	format HPRat_rat SKper10kmm_rat  
	Avg_ExperienceHP_unround Avg_ExperienceHP_round norating_f.;
run;

filename SK dde "Excel|STAR Kids-Getting Care!r3c1:r46c10" notab;

data _null_;
	set sk_all;
	file SK;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		ATC_rat '09'x WCVcomp_rat_unround  '09'x WCVcomp_rat '09'x 
		SpecTher_rat '09'x APM_survey_rat '09'x 
		Avg_Care_unround '09'x  Avg_Care_round '09'x 
	;
	format ATC_rat 
		   WCVcomp_rat_unround WCVcomp_rat 
		   SpecTher_rat APM_survey_rat  
		   Avg_Care_unround Avg_Care_round norating_f.;
run;

filename SK dde "Excel|STAR Kids-Services and Support!r3c1:r46c8" notab;

data _null_;
	set sk_all;
	file SK;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		coord_rat '09'x GNI_rat '09'x transit_rat '09'x 
		Avg_Prevention_unround  '09'x  Avg_Prevention_round '09'x 
	;
	format coord_rat GNI_rat transit_rat Avg_Prevention_unround Avg_Prevention_round norating_f.;
run;

filename SK dde "Excel|STAR Kids-Behavioral Health!r3c1:r46c8" notab;

data _null_;
	set sk_all;
	file SK;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		BHcoun_rat '09'x FUH_rat '09'x  APM_admin_rat '09'x 
		Avg_Chronic_unround  '09'x  Avg_Chronic_round '09'x 
	;
	format BHcoun_rat FUH_rat APM_admin_rat 
	Avg_Chronic_unround Avg_Chronic_round norating_f.;
run;

filename SK dde "Excel|STAR Kids-Overall!r3c1:r46c13" notab;

data _null_;
	set sk_all;
	file SK;
	put MCONAME '09'x SERVICEAREA '09'x plancode '09'x
		Avg_ExperienceHP_unround '09'x Avg_ExperienceHP_round '09'x 
		Avg_Care_unround '09'x  Avg_Care_round '09'x 
		Avg_Prevention_unround '09'x Avg_Prevention_round '09'x 
		Avg_Chronic_unround '09'x Avg_Chronic_round '09'x 
		avg_overall_unround '09'x avg_overall_round '09'x 
	;
	format Avg_ExperienceHP_round Avg_Care_round Avg_Prevention_round Avg_Chronic_round avg_overall_round  
		   Avg_ExperienceHP_unround Avg_Care_unround Avg_Prevention_unround Avg_Chronic_unround avg_overall_unround norating_f.;
run;




** ------------------------------------------ fill rating guide for RG------------------------------------------------------------;
%macro fill_rating_guide(freqdata, ratevar, address);
	filename RateGu dde "Excel|&address." notab;

	data _null_;
		set &freqdata.(where=(&ratevar. ne .));
		file RateGu;
		put &ratevar. '09'x count
		;
	run;

%mend fill_rating_guide;

%fill_rating_guide(RG_SC_1, Avg_ExperienceHP_round, 		STAR Child-ExperienceHP!r3c9:r10C10);
%fill_rating_guide(RG_SC_2, Avg_Care_round, 				STAR Child-Care!r3c10:r9C11);
%fill_rating_guide(RG_SC_3, Avg_Prevention_round, 			STAR Child-Prevention!r3c13:r10C14);
%fill_rating_guide(RG_SC_4, Avg_Chronic_round, 				STAR Child-Chronic!r3c9:r10C10);
%fill_rating_guide(RG_SC_5, avg_overall_round, 				STAR Child-Overall!r3c15:r7C16);

%fill_rating_guide(RG_SA_1, Avg_ExperienceHP_round, 		STAR Adult-ExperienceHP!r3c9:r10C10);
%fill_rating_guide(RG_SA_2, Avg_Care_round,					STAR Adult-Care!r3c10:r10C11);
%fill_rating_guide(RG_SA_3, Avg_Prevention_round, 			STAR Adult-Prevention!r3c11:r8C12);
%fill_rating_guide(RG_SA_4, Avg_Chronic_round, 				STAR Adult-Chronic!r3c11:r8C12);
%fill_rating_guide(RG_SA_5, avg_overall_round, 				STAR Adult-Overall!r3c15:r7C16);

%fill_rating_guide(RG_SP_1, Avg_ExperienceHP_round, 		STAR+PLUS-ExperienceHP!r3c9:r9C10);
%fill_rating_guide(RG_SP_2, Avg_Care_round,					STAR+PLUS-Care!r3c10:r9C11);
%fill_rating_guide(RG_SP_3, Avg_Prevention_round, 			STAR+PLUS-Prevention!r3c10:r8C11);
%fill_rating_guide(RG_SP_4, Avg_Chronic_round, 				STAR+PLUS-Chronic!r3c14:r7C15);
%fill_rating_guide(RG_SP_5, avg_overall_round, 				STAR+PLUS-Overall!r3c15:r7C16);

%fill_rating_guide(RG_SK_1, Avg_ExperienceHP_round, 		STAR Kids-ExperienceHP!r3c9:r10C10);
%fill_rating_guide(RG_SK_2, Avg_Care_round,					STAR Kids-Getting Care!r3c12:r8C13);
%fill_rating_guide(RG_SK_3, Avg_Prevention_round, 			STAR Kids-Services and Support!r3c10:r6C11);
%fill_rating_guide(RG_SK_4, Avg_Chronic_round, 				STAR Kids-Behavioral Health!r3c10:r8C11);
%fill_rating_guide(RG_SK_5, avg_overall_round, 				STAR Kids-Overall!r3c15:r6C16);

** ---- fill Missing rating guide for RG----------------------------------------------------------------------------;
proc format;
	value rating_f
		. = "No rating"
	;
run;

%macro fill_missing(freqdata, ratevar, address);
	filename RateGu dde "Excel|&address." notab;

	data _null_;
		set &freqdata.(where=(&ratevar. eq .));
		file RateGu;
		put &ratevar. '09'x count
		;
		format &ratevar. rating_f.;
	run;

%mend fill_missing;

%fill_missing(RG_SC_1, Avg_ExperienceHP_round, 	STAR Child-ExperienceHP!r11c9:r11C10);
%fill_missing(RG_SC_2, Avg_Care_round, 			STAR Child-Care!r10c10:r10C11);
%fill_missing(RG_SC_3, Avg_Prevention_round, 	STAR Child-Prevention!r11c13:r11C14);
%fill_missing(RG_SC_4, Avg_Chronic_round, 		STAR Child-Chronic!r11c9:r11C10);
%fill_missing(RG_SC_5, avg_overall_round, 		STAR Child-Overall!r8c15:r8C16);

%fill_missing(RG_SA_1, Avg_ExperienceHP_round, 	STAR Adult-ExperienceHP!r11c9:r11C10);
%fill_missing(RG_SA_2, Avg_Care_round,			STAR Adult-Care!r11c10:r11C11);
%fill_missing(RG_SA_3, Avg_Prevention_round, 	STAR Adult-Prevention!r9c11:r9C12);
%fill_missing(RG_SA_4, Avg_Chronic_round, 		STAR Adult-Chronic!r9c11:r9C12);
%fill_missing(RG_SA_5, avg_overall_round, 		STAR Adult-Overall!r8c15:r8C16);

%fill_missing(RG_SP_1, Avg_ExperienceHP_round, 	STAR+PLUS-ExperienceHP!r10c9:r10C10);
%fill_missing(RG_SP_2, Avg_Care_round,			STAR+PLUS-Care!r10c10:r10C11);
%fill_missing(RG_SP_3, Avg_Prevention_round, 	STAR+PLUS-Prevention!r9c10:r9C11);
%fill_missing(RG_SP_4, Avg_Chronic_round, 		STAR+PLUS-Chronic!r8c14:r8C15);
%fill_missing(RG_SP_5, avg_overall_round, 		STAR+PLUS-Overall!r8c15:r8C16);

%fill_missing(RG_SK_1, Avg_ExperienceHP_round, 	STAR Kids-ExperienceHP!r11c9:r11C10);
%fill_missing(RG_SK_2, Avg_Care_round,			STAR Kids-Getting Care!r9c12:r9C13);
%fill_missing(RG_SK_3, Avg_Prevention_round, 	STAR Kids-Services and Support!r7c10:r7C11);
%fill_missing(RG_SK_4, Avg_Chronic_round, 		STAR Kids-Behavioral Health!r9c10:r9C11);
%fill_missing(RG_SK_5, avg_overall_round, 		STAR Kids-Overall!r7c15:r7C16);

data _null_;
	file ddeopen;
	put '[error(false)]';
	put '[save.as("C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Output\MCO Report Cards - Composite Rating - Final.xlsx")]';
	put '[file.close(false)]';
run;





* %macro export_to_excel(filename, data_set, excel_sheet, var_list);
* 	proc format;
* 		value norating_f
* 			. = "--"
* 			other = [8.1];
* 		;
* 	run;

*     %let dsid = %sysfunc(open(&data_set));
*     %let nobs = %sysfunc(attrn(&dsid, nobs));
*     %let rc = %sysfunc(close(&dsid));
*     %let ncols = %sysfunc(countw(&var_list));
*     %let endrow = %eval(&nobs + 2); 
*     %let endcol = %eval(&ncols + 1);

*     %let dde_range = r3c1:r&endrow.c&endcol;

*     filename &filename dde "&excel_sheet.!&dde_range" notab;
    
*     data _null_;
*         set &data_set;
*         file &filename;
*         put (&var_list) ('09'x);
*         format &var_list norating_f.;
*     run;
* %mend export_to_excel;


* %export_to_excel(SC, sc_all, "Excel|STAR Child-ExperienceHP", 
*                  MCONAME SERVICEAREA plancode 
*                  HPRat_rat SCper10kmm_rat 
*                  Avg_ExperienceHP_unround Avg_ExperienceHP_round);

