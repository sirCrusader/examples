<?php
/**
 * Created by PhpStorm.
 * User: wizard
 * Date: 22.10.14
 * Time: 15:14
 */

class ReportGenerator
{
    const INTRO_FILENAME = 'intro.docx';
    const REASON_STOP_TEMPLATE = 'stop_reason.docx';
    const SUBTEST_HEAD_FILE = 'subtest_head.docx';
    const ROW_FILE_SUFFIX = '_row.docx';
    const FILE_EXTENSION = '.docx';
    const SUMMARY_FILENAME = 'summary.docx';
    const TABLE_FILENAME = 'table.docx';
    const SEMMANTICS_GRAMMAR = 'semantics_grammar.docx';
    const ONE_FILENAME = '1.docx';
    const TABLE_ONE_HEADER = 'table_1_head.docx';
    const TABLE_TWO_HEADER = 'table_2_head.docx';
    const SUMMARY_RESULTS_TABLE = 'summary_results_table.docx';
    const TEST_PROCEDURES_FILE = 'test_procedures_list.docx';
    const STATEMENTS = 'statements.docx';
    const RECOMENDATIONS = 'recomendations.docx';
    const TESTING_RESULT_FOOTER = 'testing_results_footer.docx';
    const SUMMARY_RESULT = 'summary_results.docx';
    const GFTA_LAT_INTRO_TABLE = '1_intro.docx';
    const GFTA_LAT_TABLE = '1_table.docx';
    const GFTA_LAT_ANALYSIS = '1_analysis.docx';
    const DIR_DIVIDER = '/';
    const SLI_ELIGIBILITY_TEXTS_DIR = 'sli_eligibility_checks';
    const BOLD_TITLE_STYLE = '<w:style w:type="character" w:styleId="StrongEmphasis"><w:name w:val="Strong Emphasis"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style>';
    const BOLD_HEADLINE_STYLE = '<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="Heading 4"/><w:basedOn w:val="Heading"/><w:pPr><w:spacing w:before="120" w:after="120"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:eastAsia="Droid Sans Fallback" w:cs="FreeSans"/><w:b/><w:bCs/><w:color w:val="808080"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>';

    private $reportObject, $attributes, $tests_data, $base_template_path, $base_cache_path, $temporaryFolderName, $full_templates_list, $final_docx_path, $docs_list;

    private $templates_array = array(1 => 'report_part1',
        2 => 'intro_end',
        3 => 'bahevioral_observations');

    private $excluded_comments = array('listening_comprehension_comment', 'oral_comment', 'auditory_comprehension_comment',
        'expressive_comment');

    function __construct($report)
    {
        $this->reportObject = $report;
        $this->attributes = $this->reportObject->getAttributes();
        $this->base_template_path = app_path().'/views/doctemplates';
        $this->base_cache_path =  app_path().'/storage/cache';
        $date = new DateTime();
        $this->temporaryFolderName = md5($date->getTimestamp());
        $this->docs_list = array();
    }

    public function createReport()
    {
        mkdir(app_path().'/storage/cache/'.$this->temporaryFolderName, 0777, true);
        $this->tests_data = $this->getTestList($this->reportObject->id);
        $this->full_templates_list = $this->getFullTestsList();
        $templates = $this->getTemplatePathsList();

        if (!empty($this->attributes['interviews'])) {
            array_splice($templates, 1, 0, array($this->base_template_path.'/intro_interviews.docx'));
        }

        $selected_tests_slugs = array_keys($this->tests_data);
        $tests_keys = $this->attributes['test'];
        if ('null' !== $this->attributes['test'] && !empty($this->attributes['test'])) {
            $tests_keys =  array_keys(json_decode($this->attributes['test'], true));
        }

        $generated_tests_docs = $this->generateTestsDocs($templates);

        if (count($selected_tests_slugs) > 0 || count($tests_keys) > 0) {
            $array_tests = json_decode($this->attributes['selected_tests'], true);
            $selected_tests_slugs = array_keys($array_tests);
            $test_list_doc = $this->generateTestList($selected_tests_slugs, $tests_keys);
            if (($intro_end_position = array_search(app_path().$this::DIR_DIVIDER.'storage'.$this::DIR_DIVIDER.'cache'.$this::DIR_DIVIDER.$this->temporaryFolderName.$this::DIR_DIVIDER.'general_intro_end.docx', $generated_tests_docs))) {
                array_splice($generated_tests_docs, $intro_end_position + 1, 0, array($test_list_doc));
            }
        }

        $table_info_array = json_decode($this->attributes['table_info'], true);
        if (isset($this->attributes['table_info']) && !empty($table_info_array)) {
            $summary_table_info = $this->createSummaryTable($this->attributes['table_info']);
            $array_position = array_search(app_path().$this::DIR_DIVIDER.'storage'.$this::DIR_DIVIDER.'cache'.$this::DIR_DIVIDER.$this->temporaryFolderName.$this::DIR_DIVIDER.'general_summary_results.docx', $generated_tests_docs);
            if ($array_position) {
                array_splice($generated_tests_docs, $array_position, 0, array($summary_table_info));
            }
        }

        $this->splitDocParts($generated_tests_docs, $this->attributes['report_name']);
        $this->deleteTemporaryFolder();

        return $this->final_docx_path;
    }

    /**
     * @param $report_id
     * @return array with data of the selected tests
     *
     * This function uses to get all inputed by user data for each test that was selected
     */
    private function getTestList($report_id)
    {
        $selected_tests = array();
        $tests = json_decode($this->attributes['selected_tests'], true);
        $selected_slugs = array();
        foreach ($tests as $key => $value) {
            $selected_slugs[] = $key;
        }
        $sections = Test::getSectionsArray($selected_slugs);

        $title_added = 0;
        foreach ($sections as $key => $value) {
            if ('summary' != $value) {

                if (0 == $title_added) {
                    switch($value) {
                        case 'Receptive and Expressive Language':
                            $selected_tests['language_title'] = 'section';
                            break;
                        case 'Articulation and Speech Intelligibility':
                            $selected_tests['articulation_title'] = 'section';
                            break;
                        case 'Fluency';
                            $selected_tests['fluency_title'] = 'section';
                            break;
                    }
                    $title_added = 1;
                }
                $report_test = ReportTest::findOneByReportIdAndTestSlug($report_id, $key);
                $selected_tests[$key] = json_decode($report_test->getAttributes()['test_info'], true);
                $selected_tests[$key]['section_name'] = $value;
            } else {
                $selected_tests[$key] = $value;
                $title_added = 0;
            }
        }

        return $selected_tests;
    }

    /**
     * @return array
     *
     * Function make array of the paths to all templates that will be using for generating doc-file
     */
    private function getTemplatePathsList()
    {
        $templates = array();

        foreach ($this->templates_array as $key => $value) {
            $templates[$key] = $this->base_template_path.'/'.$value.'.docx';
        }

        foreach ($this->tests_data as $key => $value) {
            if ('section' == $value || 'summary' == $value) {
                $templates[] = $this->base_template_path.$this::DIR_DIVIDER.$key.'_'.$value.'.docx';
            } else {
                $test_files = array();
                if ($handle = opendir($this->base_template_path.'/'.$key)) {
                    while (false !== ($entry = readdir($handle))) {
                        $test_files[] = $entry;
                    }
                }

                $templates[$key][] = $this->base_template_path.'/'.$key.'/'.$this::INTRO_FILENAME;
                if (($delete_key = array_search($this::INTRO_FILENAME, $test_files)) !== false) {
                    unset($test_files[$delete_key]);
                }

                $reason_template_key = array_search($this::REASON_STOP_TEMPLATE, $test_files);
                if ($reason_template_key !== false) {
                    if (isset($this->tests_data[$key]['check'][0]) && 'on' === $this->tests_data[$key]['check'][0]) {
                        $templates[$key][] = $this->base_template_path.$this::DIR_DIVIDER.$key.$this::DIR_DIVIDER.$this::REASON_STOP_TEMPLATE;
                    }
                    unset($test_files[$reason_template_key]);
                    unset($this->tests_data[$key]['check'][0]);
                }

                $subtest_head_key = array_search($this::SUBTEST_HEAD_FILE, $test_files);
                if ($subtest_head_key !== false) {
                    $templates[$key][] = $this->base_template_path.$this::DIR_DIVIDER.$key.$this::DIR_DIVIDER.$this::SUBTEST_HEAD_FILE;
                    unset($test_files[$subtest_head_key]);
                }

                if (isset($this->tests_data[$key]['check'])) {
                    switch ($key) {
                        case 'celf-p-2':
                            list($temp_array, $test_files) = $this->formatTwoTables($this->tests_data[$key]['check'], $key, $test_files, 7);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'toal-4':
                            list($temp_array, $test_files) = $this->formatTwoTables($this->tests_data[$key]['check'], $key, $test_files, 7);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'told-i-4':
                            list($temp_array, $test_files) = $this->formatTwoTables($this->tests_data[$key]['check'], $key, $test_files, 7);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'tocs':
                            list($temp_array, $test_files) = $this->formatTwoTables($this->tests_data[$key]['check'], $key, $test_files, 6);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'eft-e':
                            list($temp_array, $test_files) = $this->checkTotalScore($this->tests_data[$key]['check'], $key, $test_files);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'slsa':
                            list($temp_array, $test_files) = $this->orderSLSA($this->tests_data[$key]['check'], $key, $test_files);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'celf-4':
                            list($temp_array, $test_files) = $this->formatTableCELF($this->tests_data[$key]['check'], $key, $test_files);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);

                            list($temp_array, $test_files) = $this->orderCELF4($this->tests_data[$key]['check'], $key, $test_files);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'gfta-2':
                            list($temp_array, $test_files) = $this->orderGFTA2LAT($this->tests_data[$key]['check'], $key, $test_files);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;
                        case 'lat':
                            list($temp_array, $test_files) = $this->orderGFTA2LAT($this->tests_data[$key]['check'], $key, $test_files);
                            $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                            break;

                        default :
                            foreach ($this->tests_data[$key]['check'] as $key2 => $value2) {
                                $templates[$key][] = $this->base_template_path.'/'.$key.'/'.$key2.$this::ROW_FILE_SUFFIX;
                                if (($delete_key = array_search($key2.$this::ROW_FILE_SUFFIX, $test_files)) !== false) {
                                    unset($test_files[$delete_key]);
                                }
                            }
                            break;
                    }

                    foreach ($this->tests_data[$key]['check'] as $key2 => $value2) {
                        $templates[$key][] = $this->base_template_path.'/'.$key.'/'.$key2.'.docx';
                        if (($delete_key = array_search($key2.'.docx', $test_files)) !== false) {
                            unset($test_files[$delete_key]);
                        }
                    }
                }

                if ('fcp-r' == $key) {
                    list($temp_array, $test_files) = $this->orderFCPR($this->tests_data[$key], $key, $test_files);
                    $templates[$key] = $this->concatArrays($templates[$key], $temp_array);
                }

                for ($i =1; $i < 40; $i++) {
                    if (($delete_key = array_search($i.'.docx', $test_files)) !== false ||
                        ($delete_key = array_search($this::SUMMARY_FILENAME, $test_files)) !== false ||
                        ($delete_key = array_search('.', $test_files)) !== false ||
                        ($delete_key = array_search('..', $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }

                    $files_to_delete = array($i.$this::ROW_FILE_SUFFIX, $i . '_intro.docx', $i . '_receptive.docx', $i . '_expressive.docx',
                        $i . '_total.docx', $i . '_backward.docx', $i . '_forward.docx', $this::GFTA_LAT_INTRO_TABLE, $this::GFTA_LAT_TABLE,
                        $this::GFTA_LAT_ANALYSIS);

                    foreach ($files_to_delete as $item) {
                        if (($delete_key = array_search($item, $test_files)) !== false) {
                            unset($test_files[$delete_key]);
                        }
                    }
                }

                $test_files = $this->makeFullPath($test_files, $key);
                $templates[$key] = array_merge($templates[$key], $test_files);

                if (file_exists($this->base_template_path.'/'.$key.'/'.$this::SUMMARY_FILENAME)) {
                    $templates[$key][] = $this->base_template_path.'/'.$key.'/'.$this::SUMMARY_FILENAME;
                    if (($delete_key = array_search($this::SUMMARY_FILENAME, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }
                }
            }
        }

        $templates[] = $this->base_template_path.$this::DIR_DIVIDER.$this::SUMMARY_RESULT;
        if('null' !== $this->attributes['sli_eligibility_checks']) {
            $sli_eligibility = json_decode($this->attributes['sli_eligibility_checks'], true);

            foreach($sli_eligibility as $key => $value) {
                $templates[] = $this->base_template_path.$this::DIR_DIVIDER.$this::SLI_ELIGIBILITY_TEXTS_DIR . $this::DIR_DIVIDER . $key.'.docx';
            }
        }

        if (isset($this->attributes['impact_check']) && '0' == $this->attributes['impact_check']) {
            $templates[] = $this->base_template_path.$this::DIR_DIVIDER.$this::STATEMENTS;
        }
        if (isset($this->attributes['recommendation_check']) && '0' == $this->attributes['recommendation_check']) {
            $templates[] = $this->base_template_path.$this::DIR_DIVIDER.$this::RECOMENDATIONS;
        }
        $templates[] = $this->base_template_path.$this::DIR_DIVIDER.$this::TESTING_RESULT_FOOTER;

        return $templates;
    }

    private function formatTableCELF($test_data_array, $test_key, $test_files)
    {
        $test_templates_array = array();

        $flags = array(
            1 => false,
            2 => false,
            3 => false,
            4 => false,
        );

        for ($i = 1; $i < 16; $i++) {
            if ( isset($test_data_array[$i]) && "on" == $test_data_array[$i] ) {
                if ($i < 4 && $flags[1] == false) {
                    $test_templates_array[] = $this->base_template_path . '/' . $test_key . '/subheader_1.docx';
                    $flags[1] = true;
                }
                if ($i > 3 && $i < 9 && $flags[2] == false) {
                    $test_templates_array[] = $this->base_template_path . '/' . $test_key . '/subheader_2.docx';
                    $flags[2] = true;
                }
                if ($i > 8 && $i < 11 && $flags[3] == false) {
                    $test_templates_array[] = $this->base_template_path . '/' . $test_key . '/subheader_3.docx';
                    $flags[3] = true;

                }
                if ($i > 10 && $flags[4] == false) {
                    $test_templates_array[] = $this->base_template_path . '/' . $test_key . '/subheader_4.docx';
                    $flags[4] = true;
                }
                $test_templates_array[] = $this->base_template_path . '/' . $test_key. '/' . $i.$this::ROW_FILE_SUFFIX;

            }
            if (($delete_key = array_search($i.$this::ROW_FILE_SUFFIX, $test_files)) !== false) {
                unset($test_files[$delete_key]);
            }
        }

        for ( $i = 1; $i < 5; $i++ ) {
            if (($delete_key = array_search('subheader_'.$i . $this::FILE_EXTENSION, $test_files)) !== false) {
                unset($test_files[$delete_key]);
            }
        }

        return array($test_templates_array, $test_files);
    }

    private function orderFCPR($test_data_array, $test_key, $test_files)
    {
        $test_templates_array = array();

        foreach ( $test_data_array as $key => $value ) {
            $field_key = array_search($key.'.docx', $test_files);

            if ( !empty($value) ) {
                if ( $field_key != false ) {
                    $test_templates_array[] = $this->base_template_path . '/' . $test_key . '/' .$test_files[$field_key];
                }
            }

            unset($test_files[$field_key]);
        }

        return array($test_templates_array, $test_files);
    }

    private function orderCELF4($checks_array, $test_key, $test_files)
    {
        $test_table_templates = array();

        foreach($checks_array as $key => $value) {
            switch($key) {
                case 2:
                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . '2_intro' . $this::FILE_EXTENSION;
                    if (($delete_key = array_search('2_intro' . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }

                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . '2_receptive' . $this::FILE_EXTENSION;
                    if (($delete_key = array_search('2_receptive' . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }

                    if (isset($checks_array[6])) {
                        $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . '6_expressive' . $this::FILE_EXTENSION;
                        if (($delete_key = array_search('6_expressive' . $this::FILE_EXTENSION, $test_files)) !== false) {
                            unset($test_files[$delete_key]);
                        }
                    }

                    if (!isset($checks_array[6])) {
                        $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . '2_total' . $this::FILE_EXTENSION;
                        if (($delete_key = array_search('2_total' . $this::FILE_EXTENSION, $test_files)) !== false) {
                            unset($test_files[$delete_key]);
                        }
                    }
                    break;
                case 6:
                    if (!isset($checks_array[2])) {
                        $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . '2_intro' . $this::FILE_EXTENSION;
                        if (($delete_key = array_search('2_intro' . $this::FILE_EXTENSION, $test_files)) !== false) {
                            unset($test_files[$delete_key]);
                        }

                        $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . '6_expressive' . $this::FILE_EXTENSION;
                        if (($delete_key = array_search('6_expressive' . $this::FILE_EXTENSION, $test_files)) !== false) {
                            unset($test_files[$delete_key]);
                        }
                    }

                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . '2_total' . $this::FILE_EXTENSION;
                    if (($delete_key = array_search('2_total' . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }
                    break;
                case 9:
                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . $key . '_intro' . $this::FILE_EXTENSION;
                    if (($delete_key = array_search($key . '_intro' . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }
                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . $key . '_forward' . $this::FILE_EXTENSION;
                    if (($delete_key = array_search($key . '_forward' . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }
                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . $key . '_backward' . $this::FILE_EXTENSION;
                    if (($delete_key = array_search($key . '_backward' . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }
                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . $key . '_total' . $this::FILE_EXTENSION;
                    if (($delete_key = array_search($key . '_total' . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }
                    break;
                default:
                    $test_table_templates[] = $this->base_template_path . '/' . $test_key . '/' . $key . $this::FILE_EXTENSION;
                    if (($delete_key = array_search($key . $this::FILE_EXTENSION, $test_files)) !== false) {
                        unset($test_files[$delete_key]);
                    }
            }
        }

        return array($test_table_templates, $test_files);
    }

    private function orderSLSA($checks_array, $test_key, $test_files)
    {
        $test_table_templates = array();

        $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$this::TABLE_FILENAME;
        if (($delete_key = array_search($this::TABLE_FILENAME, $test_files)) !== false) {
            unset($test_files[$delete_key]);
        }

        $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$this::SEMMANTICS_GRAMMAR;
        if (($delete_key = array_search($this::SEMMANTICS_GRAMMAR, $test_files)) !== false) {
            unset($test_files[$delete_key]);
        }

        if (isset($checks_array[1])) {
            $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$this::ONE_FILENAME;
            if (($delete_key = array_search($this::ONE_FILENAME, $test_files)) !== false) {
                unset($test_files[$delete_key]);
            }
        }

        return array($test_table_templates, $test_files);
    }

    /**
     * Function used to make right order of the templates for tests GFTA-2 and LAT
     *
     * @param $checks_array
     * @param $test_key
     * @param $test_files
     * @return array of the needed test templates and unused test templates
     */
    private function orderGFTA2LAT($checks_array, $test_key, $test_files)
    {
        $test_table_templates = array();

        $check_keys = array_keys($checks_array);
        $exist_data = false;
        foreach ($this->tests_data[$test_key]['chart_array'] as $item) {
            if (!empty($item["'error_position'"]) ||
                !empty($item["'initial_position'"]) ||
                !empty($item["'medial_position'"]) ||
                !empty($item["'final_position'"])) {
                $exist_data = true;
                break;
            }
        }

        if (in_array(1, $check_keys)) {
            if ($exist_data) {
                $test_table_templates[] = $this->base_template_path.$this::DIR_DIVIDER.$test_key.$this::DIR_DIVIDER.$this::GFTA_LAT_INTRO_TABLE;
                $test_table_templates[] = $this->base_template_path.$this::DIR_DIVIDER.$test_key.$this::DIR_DIVIDER.$this::GFTA_LAT_TABLE;
            }

            $test_table_templates[] = $this->base_template_path.$this::DIR_DIVIDER.$test_key.$this::DIR_DIVIDER.$this::GFTA_LAT_ANALYSIS;
        }

        if (in_array(2, $checks_array)) {
            $test_table_templates[] = $this->base_template_path.$this::DIR_DIVIDER.$test_key.$this::DIR_DIVIDER.'2.docx';
        }


        return array($test_table_templates, $test_files);
    }

    private function checkTotalScore($checks_array, $test_key, $test_files)
    {
        $test_table_templates = array();
        $count = 0;
        foreach ($checks_array as $key2 => $value2) {
            $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$key2.$this::ROW_FILE_SUFFIX;
            if (($delete_key = array_search($key2.$this::ROW_FILE_SUFFIX, $test_files)) !== false) {
                unset($test_files[$delete_key]);
            }
            $count++;
        }

        if (4 == $count) {
            $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/5'.$this::ROW_FILE_SUFFIX;
            if (($delete_key = array_search('5'.$this::ROW_FILE_SUFFIX, $test_files)) !== false) {
                unset($test_files[$delete_key]);
            }
        }

        return array($test_table_templates, $test_files);
    }

    /**
     * @param $checks_array
     * @param $test_key
     * @param $test_files
     * @param $second_limit
     * @return array of templates that have right order in case test contains two tables with data
     */
    private function formatTwoTables($checks_array, $test_key, $test_files, $second_limit)
    {
        $test_table_templates = array();
        $show_first_header = $show_second_header = false;
        foreach ($checks_array as $check_key => $value) {
            if ($check_key < $second_limit) {
                $show_first_header = true;
            } else {
                $show_second_header = true;
            }
        }

        if (!$show_first_header) {
            if (($head_index = array_search($this::TABLE_ONE_HEADER, $test_files)) !== false) {
                unset($test_files[$head_index]);
            }
        }

        if (!$show_second_header) {
            if (($head_index = array_search($this::TABLE_TWO_HEADER, $test_files)) !== false) {
                unset($test_files[$head_index]);
            }
        }

        foreach ($checks_array as $key2 => $value2) {
            if ($show_first_header && $key2 < $second_limit) {
                $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$this::TABLE_ONE_HEADER;
                if (($head_index = array_search($this::TABLE_ONE_HEADER, $test_files)) !== false) {
                    unset($test_files[$head_index]);
                }
                $show_first_header = false;
            }

            if ($show_second_header && $key2 >= $second_limit) {
                $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$this::TABLE_TWO_HEADER;
                if (($head_index = array_search($this::TABLE_TWO_HEADER, $test_files)) !== false) {
                    unset($test_files[$head_index]);
                }
                $show_second_header = false;
            }

            if($key2 === 5 && $test_key === 'tocs') {
                $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$key2.'_head_row.docx';
                if (($delete_key = array_search($key2.'_head_row.docx', $test_files)) !== false) {
                    unset($test_files[$delete_key]);
                }
            }

            $test_table_templates[] = $this->base_template_path.'/'.$test_key.'/'.$key2.$this::ROW_FILE_SUFFIX;
            if (($delete_key = array_search($key2.$this::ROW_FILE_SUFFIX, $test_files)) !== false) {
                unset($test_files[$delete_key]);
            }
        }

        return array($test_table_templates, $test_files);
    }

    /**
     * @param $result_array
     * @param $temp_array
     * @return array that contains all of the elements from two arrays
     */
    private function concatArrays($result_array, $temp_array)
    {
        foreach ($temp_array as $item) {
            $result_array[] = $item;
        }

        return $result_array;
    }

    private function makeFullPath($files, $test_slug)
    {
        foreach ($files as $key => $value) {
            $files[$key] = $this->base_template_path.'/'.$test_slug.'/'.$files[$key];
        }

        return $files;
    }

    /**
     * @param $table_data
     * @return string
     * @throws \PhpOffice\PhpWord\Exception\Exception
     *
     * Function that dynamically create summary table results
     */
    private function createSummaryTable($table_data)
    {
        $data_array = json_decode($table_data, true);

        $phpWord = new \PhpOffice\PhpWord\PhpWord();

        $styleTable = array('borderColor'=>'000000',
            'borderSize'=>6,
            'cellMargin'=>50);
        $section = $phpWord->createSection();
        $section->addTextBreak();
        $section->addText('Summary of Speech and Language Testing Results',
            array(
                'size'  => 12,
                'name'  => 'Arial',
                'bold'  => true,
                'italic'    => false,
                'color' => '000000'
            )
        );
        $section->addTextBreak();

        $headTextStyle = array('bold' => true, 'name' => 'Arial');
        $cellStyle = array('borderTopSize' => 1,
            'borderTopColor' => '000000',
            'borderLeftSize' => 1,
            'borderLeftColor' => '000000',
            'borderRightSize' => 1,
            'borderRightColor' => '000000',
            'borderBottomSize' => 1,
            'borderBottomColor' => '000000');
        $cellTextStyle = array('name' => 'Arial');
        $table = $section->addTable($styleTable);
        $table->addRow(200);
        $table->addCell(3000, $cellStyle)->addText(' ');
        $table->addCell(3000, $cellStyle)->addText('Standard Score (Mean = 100)', $headTextStyle);
        $table->addCell(3000, $cellStyle)->addText('Percentile Rank (Average = 50)', $headTextStyle);

        $temporary_array_keys = array_keys($data_array);
        $temporary_array = explode("_", end($temporary_array_keys));
        $last_element = $temporary_array[0];

        for ($i = 1; $i <= $last_element; $i++) {
            $table->addRow(200);
            $table->addCell(3000, $cellStyle)->addText(!empty($data_array[$i.'_name']) ? $data_array[$i.'_name'] : ' ', $cellTextStyle);
            $table->addCell(3000, $cellStyle)->addText(!empty($data_array[$i.'_score']) ? $data_array[$i.'_score'] : ' ', $cellTextStyle);
            $table->addCell(3000, $cellStyle)->addText(!empty($data_array[$i.'_rank']) ? $data_array[$i.'_rank'] : ' ', $cellTextStyle);
        }

        /** @TODO need to fix saving of the docx-file to the other docx-parts */
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');

        $objWriter->save('/tmp/'.$this->temporaryFolderName.$this::SUMMARY_RESULTS_TABLE);
        //$objWriter->save($this->base_template_path.'/storage/cache/'.$this->temporaryFolderName.$this::SUMMARY_RESULTS_TABLE);

        return '/tmp/'.$this->temporaryFolderName.$this::SUMMARY_RESULTS_TABLE;
    }

    private function generateTestList($selected_tests_slugs, $tests_keys = null)
    {
        $test_array = array(33 => ' - Fluency Evaluation',
            34 => ' - Voice Evaluation',
            35 => ' - Interviews with',
            36 => ' - Clinical Observations',
            37 => ' - Dynamic Assessments',
            38 => ' - Alternative Assessment(s), including');
        $phpWord = new \PhpOffice\PhpWord\PhpWord();

        $list_style = array('listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER_NESTED);
        $text_style = array('name' => 'Arial');
        $section = $phpWord->createSection();
        $section->addTextBreak();
        $section->addText('Tests Administered/Test Procedures:',
            array(
                'size'  => 12,
                'name'  => 'Arial',
                'bold'  => true,
                'italic'    => false,
                'color' => '000000'
            )
        );
        $section->addTextBreak();
        foreach ($selected_tests_slugs as $slug_item) {
            $test = DB::table('tests_list')->where('slug', $slug_item)->first();
            $section->addListItem(' - '.$test->name, 0, $text_style, $list_style);
        }

        if ('null' !== $tests_keys) {
            foreach ($tests_keys as $item) {

                if (35 == $item && !empty($this->attributes['test_interview_with'])) {
                    $item_text = $test_array[$item].' '.$this->attributes['test_interview_with'];
                } elseif (38 == $item && !empty($this->attributes['test_alt_assessments'])) {
                    $item_text = $test_array[$item].' '.$this->attributes['test_alt_assessments'];
                } else {
                    $item_text = $test_array[$item];
                }
                $section->addListItem($item_text, 0, $text_style, $list_style);
            }
        }
        $section->addTextBreak();

        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');

        $objWriter->save('/tmp/'.$this->temporaryFolderName.$this::TEST_PROCEDURES_FILE);

        return '/tmp/'.$this->temporaryFolderName.$this::TEST_PROCEDURES_FILE;
    }


    /**
     * @return array
     *
     * Function gets the list of all tests that have templates and can be in the document
     */
    private function getFullTestsList()
    {
        $test_names_array = array();
        $dir = new DirectoryIterator($this->base_template_path);
        foreach ($dir as $fileinfo) {
            if ($fileinfo->isDir() && !$fileinfo->isDot()) {
                $test_names_array[] = $fileinfo->getFilename();
            }
        }

        return $test_names_array;
    }

    private function generateTestsDocs($templates)
    {
        $this->docs_list = array();

        foreach ($templates as $doc_part_key => $doc_part_value) {
            if (is_numeric($doc_part_key)) {
                if (file_exists($doc_part_value)) {
                    $this->replaceVaribles($doc_part_value, $this->attributes);
                }
            } else {
                foreach ($doc_part_value as $doc_key => $doc_value) {
                    if (file_exists($doc_value)) {
                        $this->replaceVaribles($doc_value, $this->tests_data[$doc_part_key], $doc_part_key);
                    }
                }
            }
        }

        return $this->docs_list;
    }

    private function replaceVaribles($temp_item, $data_value, $test_key = 'general'){

        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $temp_word_template = $phpWord->loadTemplate($temp_item);

        $doc_name = substr($temp_item, strrpos($temp_item, '/') + 1, strlen($temp_item));
        $this->docs_list[] = app_path().'/storage/cache/'.$this->temporaryFolderName.'/'.$test_key.'_'.$doc_name;

        switch ($temp_item) {
            case $this->base_template_path . '/slsa/'.$this::TABLE_FILENAME:
                $count = $this->countTableRows($test_key, $data_value['array']);
                $temp_word_template->cloneRow('cellValue', $count);

                $row_count = 1;
                foreach ($data_value['array'] as $value) {
                    if ($value['\'utterance\''] != "") {
                        $temp_word_template->setValue('cellName#' . $row_count, $value['\'utterance\'']);
                        $temp_word_template->setValue('cellValue#' . $row_count, $value['\'comment\'']);
                        $row_count++;
                    }
                }
                break;
            case $this->base_template_path . '/gfta-2/'.$this::GFTA_LAT_TABLE:
                $temp_word_template = $this->generateChartTable($test_key, $data_value, $temp_word_template);
                break;
            case $this->base_template_path . '/lat/'.$this::GFTA_LAT_TABLE:
                $temp_word_template = $this->generateChartTable($test_key, $data_value, $temp_word_template);
                break;

            default:
                $valiables = $temp_word_template->getVariables();
                foreach ($data_value as $field_key => $field_value) {
                    if (!is_array($field_value)) {
                        $pos = strpos($field_key, "_comment");
                        if($pos !== false && in_array($field_key, $valiables) && $field_value != "" && !in_array($field_key, $this->excluded_comments)) {
                            $temp_word_template->setValue($field_key, "");
                            $this->docs_list[] = $this->generateCommentTemplate($field_key, $field_value);
                        } else {
                            $temp_word_template->setValue($field_key, $field_value);
                        }
                    }
                }
        }

        $temp_word_template->saveAs(app_path().'/storage/cache/'.$this->temporaryFolderName.'/'.$test_key.'_'.$doc_name);
        return $this->docs_list;
    }

    private function generateChartTable($test_key, $data_value, $temp_word_template){
        $count = $this->countTableRows($test_key, $data_value['chart_array']);
        $temp_word_template->cloneRow('error_position', $count);

        $row_count = 1;
        foreach ($data_value['chart_array'] as $value) {
            if (!empty($value["'error_position'"]) || !empty($value["'initial_position'"]) || !empty($value["'medial_position'"]) || !empty($value["'final_position'"])) {
                $temp_word_template->setValue('error_position#' . $row_count, $value["'error_position'"]);
                $temp_word_template->setValue('initial_position#' . $row_count, $value["'initial_position'"]);
                $temp_word_template->setValue('medial_position#' . $row_count, $value["'medial_position'"]);
                $temp_word_template->setValue('final_position#' . $row_count, $value["'final_position'"]);
                $temp_word_template->setValue('stimulable#' . $row_count, $value["'stimulable'"]);
                $row_count++;
            }
        }

        return $temp_word_template;
    }

    private function splitDocParts($parts_array, $report_name)
    {
        $zip = new TbsZip();
        $main_content = null;
        $filename = $this->normalizeString($report_name);
        $this->final_docx_path = app_path().'/storage/cache/'.$filename.'.docx';

        $size = count($parts_array);
        for ($i = 1; $i < $size; $i++) {
            $zip->Open($parts_array[$i]);
            $content1 = $zip->FileRead('word/document.xml');
            $zip->Close();
//            $content1 = str_replace('<w:type w:val="nextPage"/>', '<w:type w:val="continuous"/>', $content1);
            //Remove section info to prevent page breaks
            $start = strpos($content1, '<w:sectPr>');
            $stop = strpos($content1, '</w:sectPr>');
            $remove_value = substr($content1, $start, $stop - $start + 11);
            $content1 = str_replace($remove_value, "", $content1);

            $p = strpos($content1, '<w:body');

            if ($p===false)
                exit("Tag <w:body> not found in document " . $parts_array[$i]);

            $p = strpos($content1, '>', $p);
            $content1 = substr($content1, $p+1);
            $p = strpos($content1, '</w:body>');

            if ($p===false)
                exit("Tag </w:body> not found in document " . $parts_array[$i]);

            $table_row_pos = strpos($parts_array[$i], '_row.docx');
            $table_subhead_row_pos = strpos($parts_array[$i], 'subheader_');

            if ($table_row_pos !== false || $table_subhead_row_pos !== false) {
                $content1 = $this->getRowXML($content1);
                $insert_pos = strrpos($main_content, '</w:tr>');
                $main_content = substr_replace($main_content, $content1, $insert_pos + 7, 0);
            } else {
                $content1 = substr($content1, 0, $p);
                $main_content .= $content1;
            }
        }

        $zip->Open($parts_array[0]);
        $content2 = $zip->FileRead('word/document.xml');
        $p = strpos($content2, '</w:body>');

        if ($p===false)
            exit("Tag </w:body> not found in document " . $parts_array[0]);

        $content2 = substr_replace($content2, $main_content, $p, 0);
        $zip->FileReplace('word/document.xml', $content2, TBSZIP_STRING);

        $this->updateFinalDocStyles($zip, $parts_array[0]);

// Save the merge into a third file
        $zip->Flush(TBSZIP_FILE, $this->final_docx_path);
    }

    private function getRowXML($content1)
    {
        $table_content = null;
        $flag = true;
        $pos = 0;
        while($flag) {
            $start_row = strpos($content1, '<w:tr>');
            $end_row = strpos($content1, '</w:tr>');
            $row_content = substr($content1, $start_row, $end_row - $start_row + 7);
            if ($start_row >= $pos) {
                $table_content .= $row_content;
                $content1 = substr_replace($content1, "", $start_row, $end_row - $start_row + 7);
                $pos = $start_row;
            } else {
                $flag = false;
            }
        }

        return $table_content;
    }

    /** Function used to create safe filename */
    private function normalizeString($str = '')
    {
        $str = strip_tags($str);
        $str = preg_replace('/[\r\n\t ]+/', ' ', $str);
        $str = preg_replace('/[\"\*\/\:\<\>\?\'\|]+/', ' ', $str);
        $str = strtolower($str);
        $str = html_entity_decode( $str, ENT_QUOTES, "utf-8" );
        $str = htmlentities($str, ENT_QUOTES, "utf-8");
        $str = preg_replace("/(&)([a-z])([a-z]+;)/i", '$2', $str);
        $str = str_replace(' ', '-', $str);
        $str = rawurlencode($str);
        $str = str_replace('%', '-', $str);

        return $str;
    }

    /**
     * This function delete temporary directory with all parts of the generated doc
     */
    private function deleteTemporaryFolder()
    {
        $dir_path = app_path().'/storage/cache/'.$this->temporaryFolderName;
        if (! is_dir($dir_path)) {
            throw new InvalidArgumentException("$dir_path must be a directory");
        }
        if (substr($dir_path, strlen($dir_path) - 1, 1) != '/') {
            $dir_path .= '/';
        }
        $files = glob($dir_path . '*', GLOB_MARK);
        foreach ($files as $file) {
            if (is_dir($file)) {
                self::deleteDir($file);
            } else {
                unlink($file);
            }
        }
        rmdir($dir_path);

        $table_info_array = json_decode($this->attributes['table_info'], true);
        if (isset($this->attributes['table_info']) && !empty($table_info_array)) {
            unlink('/tmp/' . $this->temporaryFolderName . $this::SUMMARY_RESULTS_TABLE);
        }

        unlink('/tmp/'.$this->temporaryFolderName.$this::TEST_PROCEDURES_FILE);
    }

    private function generateCommentTemplate($value, $text) {
        $path = $this->base_cache_path . '/' . $this->temporaryFolderName . '/' . $value . '.docx';

        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->createSection();
        $section->addText($text,  array(
            'size'  => 12,
            'name'  => 'Arial',
        ));
        $section->addTextBreak();

        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save($path);

        return $path;
    }

    private function countTableRows($key, $array)
    {
        $count = 0;

        switch ($key) {
            case 'slsa':
                foreach ($array as $key => $value) {
                    if (($value['\'utterance\''] != "") || ($value['\'comment\''] != "")) {
                        $count++;
                    }
                }
                break;
            case 'gfta-2':
                foreach ($array as $key => $value) {
                    if (!empty($value["'error_position'"]) || !empty($value["'initial_position'"]) || !empty($value["'medial_position'"]) || !empty($value["'final_position'"])) {
                        $count++;
                    }
                }
                break;
            case 'lat':
                foreach ($array as $key => $value) {
                    if (!empty($value["'error_position'"]) || !empty($value["'initial_position'"]) || !empty($value["'medial_position'"]) || !empty($value["'final_position'"])) {
                        $count++;
                    }
                }
                break;
        }

        return $count;
    }

    private function updateFinalDocStyles($zip, $file)
    {
        $styles_content = $this::BOLD_TITLE_STYLE . $this::BOLD_HEADLINE_STYLE;

        $styles = $zip->FileRead('word/styles.xml');
        $style = strpos($styles, '</w:styles>');

        if ($style===false)
            exit("Tag </w:styles> not found in document " . $file);

        $styles = substr_replace($styles, $styles_content, $style, 0);

        $zip->FileReplace('word/styles.xml', $styles, TBSZIP_STRING);
    }
}