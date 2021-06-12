<?php
/*
 *             ********************
 *                  sam_victory
 *              ******************
 */

if (!isset($argv[1]) || !is_file($argv[1])) {
    exit("[usage] ./$argv[0] myexcelfile.xlsx\n");
}

$handle = fopen("zip://$argv[1]#xl/sharedStrings.xml", 'r');
if ($handle) {
    $result = '';
    while (!feof($handle)) {
        $result .= fread($handle, 8192);
    }
    //echo $result;

    fclose($handle);
    $content = str_replace("\n", "", $result);
    if (preg_match_all('/\<si>(.*?)\<\/si>/s', $result, $match)) {
        $dico = $match[1];
        foreach ($match[1] as $k => $v) {
            $dico[$k] = strip_tags($v);
        }

    }
}

if (($zip = zip_open($argv[1]))) {
    while ($zip_entry = zip_read($zip)) {
        if (!zip_entry_open($zip, $zip_entry)) {
            continue;
        }

        $filename = zip_entry_name($zip_entry);
        if (preg_match("/xl\/worksheets\/sheet([0-9]{0,3}).xml/", $filename, $fileId)) {
            //     Open File
            $filesize = intval(zip_entry_filesize($zip_entry));
            $content = zip_entry_read($zip_entry, $filesize);
            //echo "\n\n\n\nFilename :; $filename;\n\n\n\n";
            $rowReg = '/\<row [A-Za-z-0-9^_.:="\' ]{0,200}>(.*?)\<\/row>/s';
            if (preg_match_all($rowReg, $content, $data)) {
                $rows = $data[1];
                //       Read All Data Line
                foreach ($rows as $k => $v) {
                    $colReg = '/\<c [A-Za-z-0-9^_.:="\' ]{0,200}(\>\<v>(.*?)\<\/v>|\/\>)/s';
                    if (preg_match_all($colReg, $v, $row)) {
                        $rowData = $row[2];
                        //       Extract All Record from Line
                        foreach ($rowData as $k => $v) {
                            if ($k != count($rowData) - 1) {
                                if (strstr($row[0][$k], 't="s"')) {
                                    $v = str_replace("\n", '\n', $dico[$v]);
                                }
                                // mettre "," si vous voulez que le séparateur soit un ",".
                                // si vous voulez ";" comme séparateur.
                                // il suffit de changer le ',' => ';'
                                echo '' . addslashes(($v == '/>') ? '' : $v) . ',';
                            } else {
                                echo '' . addslashes(($v == '/>') ? '' : $v) . '';
                            }

                        }
                        echo "\n";
                    }
                }
            }
        }
        zip_entry_close($zip_entry);
    }
    zip_close($zip);
}
