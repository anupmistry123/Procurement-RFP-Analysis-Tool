CREATE OR REPLACE AGGREGATE FUNCTION __utils_median_num (v_metric_col decimal(17,4), v_out_col decimal(17,4), v_cond_col TINYINT(1))
RETURNS decimal(17,4)
BEGIN
## this function takes the median value from a table using the "GROUP BY" clause.
## it checks if there are odd values or even values and performs the correct calculation. 
    DECLARE v_metric_str text DEFAULT NULL; DECLARE v_out_str text DEFAULT NULL;
    DECLARE v_list_length int DEFAULT 0; DECLARE v_list_counter int DEFAULT 0;
    DECLARE v_thr_num decimal(17,4); DECLARE v_matched bool DEFAULT FALSE;
    DECLARE CONTINUE HANDLER FOR NOT FOUND
    BEGIN
        SET v_list_length := (length(v_out_str) - length(replace(v_out_str, '|', ''))) / length('|') + 1;
        RETURN 
            ifnull(CASE v_list_length MOD 2
                WHEN 0 THEN (substring_index(substring_index(v_out_str, '|', floor(v_list_length / 2)), '|', -1) + substring_index(substring_index(v_out_str, '|', floor(v_list_length / 2) + 1), '|', -1)) / 2
                WHEN 1 THEN substring_index(substring_index(v_out_str, '|', floor(v_list_length / 2) + 1), '|', -1)
            END, 0);
    END;
    LOOP
        FETCH GROUP NEXT ROW;
        IF v_cond_col THEN
            SET v_list_length := (length(v_metric_str) - length(replace(v_metric_str, '|', ''))) / length('|') + 1;
            IF v_metric_str IS NULL THEN
                SET v_metric_str := v_metric_col; SET v_out_str := v_out_col;
            ELSE
                SET v_matched := FALSE;
                FOR v_list_counter IN 1 .. v_list_length DO
                    IF NOT v_matched AND v_list_counter = 1 AND v_metric_col <= substring_index(v_metric_str, '|', 1) / 1 THEN
                        SET v_metric_str := concat(v_metric_col, '|', v_metric_str);
                        SET v_out_str := concat(v_out_col, '|', v_out_str);
                        SET v_matched := TRUE;
                    ELSEIF NOT v_matched AND v_list_counter = v_list_length AND v_metric_col > substring_index(substring_index(v_metric_str, '|', v_list_length), '|', -1) THEN
                        SET v_metric_str := concat(v_metric_str, '|', v_metric_col);
                        SET v_out_str := concat(v_out_str, '|', v_out_col);
                        SET v_matched := TRUE;
                    ELSEIF NOT v_matched AND v_metric_col <= substring_index(substring_index(v_metric_str, '|', v_list_counter), '|', -1) THEN
                        SET v_metric_str := concat(substring_index(v_metric_str, '|', v_list_counter - 1), '|',  v_metric_col, '|', substring_index(v_metric_str, '|', v_list_counter + 1 - v_list_length));
                        SET v_out_str := concat(substring_index(v_out_str, '|', v_list_counter - 1), '|',  v_out_col, '|', substring_index(v_out_str, '|', v_list_counter - 1 - v_list_length));
                        SET v_matched := TRUE;                        
                    END IF;
                END FOR;
            END IF;
        END IF;
    END LOOP;
END
