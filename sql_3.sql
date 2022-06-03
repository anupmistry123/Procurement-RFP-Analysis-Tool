CREATE OR REPLACE PROCEDURE __output_dynamic_controls (
 IN `v_db_table` TEXT
)

BEGIN
## this procedure is an "output" procedure that takes a parameter (a reference to a database table) and outputs it into the analysis tool with easy to read columns.
## in many cases there are dynamic columns so the output code is dynamiclly created. I.e. in some tools, 3 columns appear, in others, only 2. 
    CASE
      WHEN v_db_table = '__controls_basket_levels' THEN 
      SELECT __config_max_basket_level() INTO @v_max_level;
      WITH RECURSIVE num_gen AS (SELECT 1 num UNION ALL SELECT num + 1 FROM num_gen WHERE num < @v_max_level)
                             , labels AS (
                                         SELECT if(num > @v_max_level, NULL, concat('level_', num)) b_level
                                              , num ordinal
                                           FROM num_gen
                                         )
      SELECT concat('

SELECT sheet_name AS `Sheet Name`
     , sheet_id AS `Sheet Position` 
     , basket_id AS `Basket ID`
     , basket_urn AS `Basket URN`
     ,', (SELECT group_concat(concat(b_level, ' AS `', __utils_cap_first(b_level), '`') SEPARATOR ',') FROM labels), '
     , basket_alias AS `Basket Alias`     
     , use_for_outliers AS `Use Basket For Outlier Parameters`
     , use_for_savings_scenarios AS `Use Basket For Savings Scenarios`
     , use_for_coverage_heatmap AS `Use Basket For Coverage Heatmap`
  FROM __controls_basket_levels
                   ')
        INTO @v_qry; 
      WHEN v_db_table = '__controls_basket_coverage' THEN
      SET @v_qry := concat('
SELECT sheet_name        AS `Sheet Name`
     , sheet_id          AS `Sheet Position`
     , basket_id AS `Basket ID`
     , basket_description AS `Basket Description`'
     , __outputs_qrybuilder('dynamic-levels-cols-baskets-only', 'bc'), '
     , good_coverage_threshold AS `Good Coverage Threshold`
  FROM __controls_basket_coverage bc
;'
                          )
    ;
    WHEN v_db_table = '__controls_bidders' THEN
      SET @v_qry:= concat('
SELECT ', if(__config_use_bidder_code(), 'code           AS `Bidder Code`
     , ', ''), 'name           AS `Bidder Name`
     , excl_scenarios AS `Exclude From Award Scenarios`
     , ', if(!__config_use_custom_baseline(), 'baseline       AS `Baseline Scenario`
     , ', ''), (SELECT group_concat(concat(shortlist_name, ' AS `', shortlist_alias,'`')) FROM __config_shortlists), '
  FROM __controls_bidders
;'
                         )
    ;
      WHEN v_db_table = '__controls_outlier_parameters' THEN
        SET @v_qry := concat('
SELECT sheet_name        AS `Sheet Name`
     , sheet_id          AS `Sheet Position`
     , basket_id AS `Basket ID`
     , basket_description AS `Basket Description`'
     , __outputs_qrybuilder('dynamic-levels-cols-baskets-only', 'op'), '
     , outlier_lower_bound AS `Outlier -`
     , outlier_upper_bound AS `Outlier +`
  FROM __controls_outlier_parameters op
;'
                          )
      ;
      WHEN v_db_table = '__controls_savings_scenarios' THEN  
      SELECT __config_current_max_basket_level(), __config_preferred_sup_sel_num()
      INTO @v_max_level               , @v_pss_num
      ;
      WITH RECURSIVE num_gen AS (SELECT 1 num UNION ALL SELECT num + 1 FROM num_gen WHERE num < if(@v_pss_num > @v_max_level, @v_pss_num, @v_max_level))
                                      , labels AS (
                                                  SELECT if(num > @v_pss_num, NULL, concat('preferred_supplier_', char(96 + num))) ps_label
                                                       , if(num > @v_max_level, NULL, concat('level_', num)) b_level
                                                       , num ordinal
                                                    FROM num_gen
                                                  )
     SELECT concat('
  SELECT sheet_name          AS `Sheet Name`
       , sheet_id            AS `Sheet Position`
       , basket_id AS `Basket ID`
       , basket_description AS `Basket Description`' , ifnull((SELECT group_concat(concat('
       , ', b_level, ' AS `Basket Level ', ordinal, '`') SEPARATOR '') FROM labels), ''), '
       , savings_threshold   AS `Savings Threshold`
       , top_n_selection_num AS `Top N`
       , ', (SELECT group_concat(concat(shortlist_name, ' AS `', shortlist_alias, '`') SEPARATOR '
       , ')    FROM __config_shortlists), '
       , ', (SELECT group_concat(concat(ps_label, ' AS `', __utils_cap_first(ps_label), '`') SEPARATOR '
       , ')    FROM labels), '
    FROM __controls_savings_scenarios
  ;'          
                   )
       INTO @v_qry;
      ELSE
        SET @v_qry = 'SELECT ''ERROR'' AS `Error Flag`, ''Incorrect input parameter, please resolve.'' AS `Error Detail`';
    END CASE;
    EXECUTE IMMEDIATE @v_qry;
END
