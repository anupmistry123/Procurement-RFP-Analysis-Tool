CREATE OR REPLACE PROCEDURE __stage_raw ()

BEGIN
    #################################################################################################################################
    ## Stage raw data
    ##
    ## This procedure moves data out of the raw table inputs into a selection of staged data tables, including a normalised table for all pricing.
    ## For each data sheet_id, we get the following tables:
    ##     - if there are item_data columns (fixed data per supplier), we will get tables named __responses_items{_i} in the (e.g. __responses_items_1)
    ##     - if there are additional_info columns  (data which is different per supplier, not pricing), we will get tables named  __responses_info{_i} in the responses schema
    ## We will also get the table prices in the public schema, which is normalised.
    ## This is acheived by joining the frame onto the raw data tables.
    ##
    ## log the procedure
    CALL __log_new_proc('stage raw data', @v_proc_id);
    ## update raw config with the max character length for each of the columns
    FOR v_record IN (SELECT * FROM __config_raw) DO
        EXECUTE IMMEDIATE concat('SELECT max(length(`', v_record.cleaned_column_name, '`)) INTO @v_max_char FROM `__raw_', lower(v_record.sheet_id), '`');
        UPDATE __config_raw
           SET max_column_length = if(@v_max_char = 0, 1, @v_max_char)
         WHERE sheet_id = v_record.sheet_id
           AND cleaned_column_name = v_record.cleaned_column_name
        ;
    END FOR;
    UPDATE __config_raw SET data_type = 'char' WHERE data_type = '';
    ## create the prices table
    CREATE OR REPLACE TABLE _prices (
        frame_id            int
      , class_id            int
      , import_id           int
      , sheet_id            int
      , urn_id              int
      , bidder_name         varchar(120)
      , price_type_id       int
      , price               decimal(17,4)
      , blend_price_type_id int
    );
    ## create the urn mapping table
    SELECT group_concat(concat('
    SELECT DISTINCT ', sheet_id , ' AS sheet_id
         , CAST(`',cleaned_column_name, '` AS CHAR(150)) AS urn
      FROM __raw_', sheet_id, '
     WHERE import_id = (SELECT max(id) FROM __log_import where error_message IS NULL)'
                 ) ORDER BY sheet_id SEPARATOR '
     UNION ALL') v_stmt
      INTO @v_stmt
      FROM __config_raw
     WHERE special_name = 'urn'
    ;
    EXECUTE IMMEDIATE concat('CREATE OR REPLACE TABLE _temp_um ', @v_stmt);
    ALTER TABLE _temp_um ADD COLUMN id INT AUTO_INCREMENT KEY FIRST;
    CREATE OR REPLACE TABLE _urn_mapping SELECT * FROM _temp_um ORDER BY id;
    CREATE INDEX _urn_mapping_sheet_id ON _urn_mapping(sheet_id);
    CREATE INDEX _urn_mapping_id ON _urn_mapping(id);
    CREATE INDEX _urn_mapping_comp ON _urn_mapping(sheet_id, urn);
    DROP TABLE IF EXISTS _temp_um;
    ## Start the main process, by entering a cursor loop through the sheets. Rebates are handled differently, at the end of this process
    SET @err_str := '';
    sheet_ids: FOR v_codes IN (
                              SELECT r.sheet_id
                                   , t.sheet_name
                                   , sum(special_name = 'bidder') > 0                    AS bid_check
                                   , sum(special_name = 'item_quantity') > 0             AS qu_check
                                   , sum(special_name = 'urn') > 0                       AS urn_check
                                   , sum(special_name = 'line_item_incumbent') > 0       AS line_item_inc_check
                                   , sum(column_type = 'item_data') > 0                  AS id_check
                                   , sum(column_type = 'price_column') > 0               AS cc_check
                                   , sum(column_type = 'additional_info') > 0            AS ai_check
                                FROM __config_raw r
                                JOIN __config_sheets t ON 1 = 1
                                 AND t.id = r.sheet_id
                               WHERE t.rebate_sheet = ''
                               GROUP BY 1, 2
                              ) DO
        ## extraction only occurs when there are price columns against the tab code. 
        IF v_codes.cc_check THEN
            ## check that there is a urn column.
            IF NOT v_codes.urn_check THEN
                SET @err_str := concat('pricing sheet ', v_codes.sheet_name, ' has no URN column flagged in raw config. Please flag a URN column to proceed');
                LEAVE sheet_ids;
            ELSE
                ## get the urn column
                SELECT cleaned_column_name
                  INTO @v_urn_col
                  FROM __config_raw
                 WHERE sheet_id = v_codes.sheet_id
                   AND special_name = 'urn'
                ;
                SET @v_qry_head := concat('
SELECT f.id frame_id
     , f.class_id
     , r.import_id
     , ', v_codes.sheet_id, ' AS sheet_id
     , u.id AS urn_id
     , f.bidder_name'
                                        );
                ## get the bidder column if required.
                IF __config_eflow_tender() THEN
                    IF NOT v_codes.bid_check THEN
                        SET @err_str := concat('pricing sheet ', v_codes.sheet_name, ' has no Bidder column flagged in raw config, despite being flagged as an eFlow tender. Either switch to non-Eflow in env_vars or flag the bidder column in raw config.');
                        LEAVE sheet_ids;
                    ELSE
                        SELECT concat('CAST(`', cleaned_column_name, '` AS '
                                    , CASE data_type
                                        WHEN 'char'     THEN concat('char(', max_column_length,')')
                                        WHEN 'numeric'  THEN 'decimal(17,4)'
                                        WHEN 'monetary' THEN 'decimal(17,4)'
                                      END, ') AS bidder')
                             , cleaned_column_name
                          INTO @v_cast_bidder_col
                             , @v_bidder_col
                          FROM __config_raw
                         WHERE sheet_id = v_codes.sheet_id
                           AND special_name = 'bidder'
                        ;
                    END IF;
                END IF;
                SET @v_qry_tail := concat('
  FROM `__raw_', v_codes.sheet_id, '` r
  JOIN _frame f ON 1 = 1
   AND f.current_import_id = r.import_id', if(__config_eflow_tender(), concat('
   AND f.bidder_name = r.', @v_bidder_col), ''), '
  JOIN _urn_mapping u ON 1 = 1
   AND u.urn = r.`', @v_urn_col,'`
   AND u.sheet_id = ', v_codes.sheet_id, '
');
                IF v_codes.id_check THEN
                    SELECT group_concat(concat('     , CAST(', if(data_type = 'numeric', 'CAST(', '') ,'nullif(`', cleaned_column_name, '`, '''') AS '
                                             , if(special_name <> 'item_quantity',
                                                   CASE data_type
                                                     WHEN 'char'     THEN concat('char(', max_column_length,')')
                                                     WHEN 'numeric'  THEN 'decimal(17,4)'
                                                     WHEN 'monetary' THEN 'decimal(17,4)'
                                                   END
                                                 , 'decimal(17,4)'
                                                 ) , if(data_type = 'numeric', ') AS int', ''), ') AS `', if(mapped_column_name = '', cleaned_column_name, mapped_column_name), '`', char(10)) SEPARATOR '')
                      INTO @v_id_col_str                  
                      FROM __config_raw
                     WHERE sheet_id = v_codes.sheet_id
                       AND column_type = 'item_data'
                       AND special_name <> 'urn'
                    ;
                    SET @v_stmt := concat('CREATE OR REPLACE TABLE `__responses_items_', v_codes.sheet_id, '` AS 
SELECT DISTINCT
       f.class_id
     , ', v_codes.sheet_id, ' AS sheet_id
     , u.id AS urn_id'
     , @v_id_col_str
     , @v_qry_tail
                                         );
                    CALL __log_dynamic_code(@v_proc_id, concat(v_codes.sheet_name, ' items creation'), @v_stmt);
                    EXECUTE IMMEDIATE @v_stmt;
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_items_', v_codes.sheet_id, '_class_id` ON `__responses_items_', v_codes.sheet_id, '`(class_id);');
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_items_', v_codes.sheet_id, '_sheet_id` ON `__responses_items_', v_codes.sheet_id, '`(sheet_id);');
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_items_', v_codes.sheet_id, '_comp_1` ON `__responses_items_', v_codes.sheet_id, '`(class_id, sheet_id, urn_id);');
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_items_', v_codes.sheet_id, '_comp_2` ON `__responses_items_', v_codes.sheet_id, '`(sheet_id, urn_id);');
                END IF;
                IF v_codes.ai_check THEN
                    SELECT group_concat(concat('     , CAST(', if(data_type = 'numeric', 'CAST(', '') ,'nullif(`', cleaned_column_name, '`, '''') AS '
                                             , CASE data_type
                                                 WHEN 'char'     THEN concat('char(', max_column_length,')')
                                                 WHEN 'numeric'  THEN 'decimal(17,4)'
                                                 WHEN 'monetary' THEN 'decimal(17,4)'
                                               END, ')', if(data_type = 'numeric', ' AS int)', ''),' AS `', if(mapped_column_name = '', cleaned_column_name, mapped_column_name), '`', char(10)) SEPARATOR '')
                      INTO @v_ai_col_str                  
                      FROM __config_raw
                     WHERE sheet_id = v_codes.sheet_id
                       AND column_type = 'additional_info'
                    ;
                SET @v_qry_head := concat('
SELECT f.id frame_id
     , f.class_id
     , r.import_id
     , ', v_codes.sheet_id, ' AS sheet_id
     , u.id AS urn_id
     , f.bidder_name'
                                        );
                    SET @v_stmt := concat('CREATE OR REPLACE TABLE `__responses_info_', v_codes.sheet_id, '` AS '
                                        , @v_qry_head, @v_ai_col_str, @v_qry_tail);
                    CALL __log_dynamic_code(@v_proc_id, concat(v_codes.sheet_id, ' info creation'), @v_stmt);
                    EXECUTE IMMEDIATE @v_stmt;
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_info_', v_codes.sheet_id, '_sheet_id` ON `__responses_info_', v_codes.sheet_id, '`(sheet_id);');
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_info_', v_codes.sheet_id, '_frame_id` ON `__responses_info_', v_codes.sheet_id, '`(frame_id);');
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_info_', v_codes.sheet_id, '_class_id` ON `__responses_info_', v_codes.sheet_id, '`(class_id);');
                    EXECUTE IMMEDIATE concat('CREATE INDEX `__responses_info_', v_codes.sheet_id, '_comp` ON `__responses_info_', v_codes.sheet_id, '`(frame_id, sheet_id, urn_id);');
                END IF;
                SET @v_stmt := '';
                FOR v_price_cols IN (SELECT p.id, if(r.mapped_column_name = '', r.cleaned_column_name, r.mapped_column_name) db_column_name FROM __config_prices p JOIN __config_raw r ON r.original_column_name = p.price_type AND r.sheet_id = p.sheet_id WHERE p.sheet_id = v_codes.sheet_id AND r.column_type = 'price_column') DO
                    SET @v_stmt := concat(@v_stmt, if(@v_stmt = '', '', concat(char(10), ' UNION ALL ')), @v_qry_head, '
     , CAST(''', v_price_cols.id, ''' AS int) AS price_type_id
     , CAST(ifnull(nullif(`', v_price_cols.db_column_name, '`, ''''), 0) AS decimal(17,4)) AS price
     , NULL', @v_qry_tail);
                END FOR;
                CALL __log_dynamic_code(@v_proc_id, concat(v_codes.sheet_id, ' insertion into _prices'), concat('INSERT INTO _prices ', char(10), @v_stmt));
                EXECUTE IMMEDIATE concat('INSERT INTO _prices ', @v_stmt);
            END IF;
        END IF;
    END FOR sheet_ids;
    IF !(@err_str = '') THEN SIGNAL SQLSTATE '45000' SET MESSAGE_TEXT = @err_str; END IF;
    CREATE INDEX _prices_frame_id     ON _prices(frame_id);
    CREATE INDEX _prices_class_id     ON _prices(class_id);
    CREATE INDEX _prices_price_type_id ON _prices(price_type_id);
    CREATE INDEX _prices_comp_items   ON _prices(class_id, sheet_id, urn_id);
    CREATE INDEX _prices_sheet_urn ON _prices(sheet_id, urn_id);
    CREATE INDEX _prices_price_type   ON _prices(sheet_id, price_type_id);
    CREATE INDEX _prices_comp_info    ON _prices(frame_id, sheet_id, urn_id);
    CREATE INDEX _prices_comp_full    ON _prices(frame_id, sheet_id, urn_id, price_type_id);
    CREATE INDEX _prices_comp_analyze ON _prices(class_id, sheet_id, urn_id, price_type_id);
     ## create the rebates table
    CREATE OR REPLACE TABLE _rebates (
          frame_id             bigint
        , class_id             bigint
        , import_id            bigint
        , bidder_name          varchar(40)
        , sheet_id             int
        , spend_from           decimal(17,4)
        , spend_to             decimal(17,4)
        , percentage_discount  decimal(17,4)
      );     
    IF (SELECT (count(*) = 1) FROM __config_sheets WHERE rebate_sheet = 'Y') THEN
        SELECT ifnull(concat(',', group_concat(concat('NULLIF(r.`', r.cleaned_column_name, '`, '''') AS `', r.special_name,'`') 
                                               ORDER BY CASE r.special_name 
                                                          WHEN 'spend_from' THEN 0
                                                          WHEN 'spend_to' THEN 1
                                                          WHEN 'percentage_discount' THEN 2
                                                        END)), '')
          INTO @v_stmt
          FROM __config_raw r
          JOIN __config_sheets t ON 1 = 1
           AND t.id = r.sheet_id
         WHERE 1 = 1
           AND r.special_name REGEXP 'spend_(from|to)|percentage_discount'
           AND t.rebate_sheet = 'Y'
        ;
        SELECT id INTO @v_rebates_sheet_id FROM __config_sheets WHERE rebate_sheet = 'Y';
        SET @v_stmt := concat('
SELECT f.id frame_id
     , f.class_id
     , r.import_id
     , f.bidder_name
     , ', @v_rebates_sheet_id, ' AS sheet_id'
     , @v_stmt,'
FROM `__raw_', @v_rebates_sheet_id, '` r
JOIN _frame f ON 1 = 1
 AND f.current_import_id = r.import_id
 AND f.baseline = ''''', if(__config_eflow_tender(), concat('
 AND f.bidder_name = r.', @v_bidder_col), '')
                             );
        CALL __log_dynamic_code(@v_proc_id, ' insertion into _rebates', @v_stmt);                 
        EXECUTE IMMEDIATE concat('INSERT INTO _rebates ', @v_stmt);
    ELSE
         UPDATE __ux_journey_sheets
           SET active = ''
         WHERE sheet_name_override = 'Rebates'
        ;
    END IF;
    CREATE INDEX _rebates_frame_id  ON _rebates(frame_id);
    CREATE INDEX _rebates_class_id  ON _rebates(class_id);
    CREATE INDEX _rebates_comp_full ON _rebates(frame_id, sheet_id, spend_from, spend_to);
END
