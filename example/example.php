<?php

use arsenweb\Reader\ExtReader;
use arsenweb\Writer\ExtWriter;

require_once './vendor/autoload.php';

$reader = ExtReader::read('./example/excel.xlsx')
    ->setMaskList(
        [
            '${HEADER}',
            '${PERIOD}',
            '${TOTAL_VALUE_1}',
            '${TOTAL_VALUE_2}',
            '${TOTAL_VALUE_3}',
            '${TOTAL_VALUE_4}',
            '${TOTAL_VALUE_5}',
            '${TOTAL_VALUE_6}',
            '${ITEMS}',
        ]
    );

$writer = ExtWriter::setReader($reader)
    ->setString('${HEADER}', 'HEADER')
    ->setString('${PERIOD}', 'With from 01.01.2021 to 01.02.2021')
    ->setInt('${TOTAL_VALUE_1}', 1)
    ->setInt('${TOTAL_VALUE_2}', 22)
    ->setInt('${TOTAL_VALUE_3}', 333)
    ->setFloat('${TOTAL_VALUE_4}', 4.444, 3)
    ->setFloat('${TOTAL_VALUE_5}', 55.555, 3)
    ->setFloat('${TOTAL_VALUE_6}', 666.666, 3)
    ->setAssociativeArray(
        '${ITEMS}',
        [
            [1, 2, 3, 4, 5, 6],
            [2, 3, 4, 5, 6, 7],
            [3, 4, 5, 6, 7, 8],
            [4, 5, 6, 7, 8, 9],
            [5, 6, 7, 8, 9, 10],
            [6, 7, 8, 9, 10, 11],
        ],
        function($items) {
            $result = [];
            foreach($items as $item) {
                $item[] = array_sum($item);
                $result[] = $item;
            }

            return $result;
        })
    ->save('./example/excel_result.xlsx');

ExtWriter::setReader($reader)
    ->setString('${HEADER}', 'HEADER')
    ->setString('${PERIOD}', 'With from 01.01.2021 to 01.02.2021')
    ->setArray(
        '${TOTAL_VALUE_1}',
        [
            1,
            22,
            333,
            4.444,
            55.555,
            666.666,
        ],
        null,
        ExtWriter::DIRECTION_RIGHT
    )
    ->setAssociativeArray(
        '${ITEMS}',
        [
            [1, 2, 3, 4, 5, 6],
            [2, 3, 4, 5, 6, 7],
            [3, 4, 5, 6, 7, 8],
            [4, 5, 6, 7, 8, 9],
            [5, 6, 7, 8, 9, 10],
            [6, 7, 8, 9, 10, 11],
        ],
        function($items) {
            $result = [];
            foreach($items as $item) {
                $item[] = array_sum($item);
                $result[] = $item;
            }

            return $result;
        })
    ->save('./example/excel_result_2.xlsx');
