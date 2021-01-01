<?php

namespace arsenweb\Helpers;

use arsenweb\Exceptions\ColumnException;

class ColumnHelper implements IColumnHelper
{
    /** @var array */
    protected $sourse = [
        'A' => 1, 'B' => 2, 'C' => 3,
        'D' => 4, 'E' => 5, 'F' => 6,
        'G' => 7, 'H' => 8, 'I' => 9,
        'J' => 10, 'K' => 11, 'L' => 12,
        'M' => 13, 'N' => 14, 'O' => 15,
        'P' => 16, 'Q' => 17, 'R' => 18,
        'S' => 19, 'T' => 20, 'U' => 21,
        'V' => 22, 'W' => 23, 'X' => 24,
        'Y' => 25, 'Z' => 26, 'a' => 1,
        'b' => 2, 'c' => 3, 'd' => 4,
        'e' => 5, 'f' => 6, 'g' => 7,
        'h' => 8, 'i' => 9, 'j' => 10,
        'k' => 11, 'l' => 12, 'm' => 13,
        'n' => 14, 'o' => 15, 'p' => 16,
        'q' => 17, 'r' => 18, 's' => 19,
        't' => 20, 'u' => 21, 'v' => 22,
        'w' => 23, 'x' => 24, 'y' => 25,
        'z' => 26,
    ];

    /**
     * @param string $colStr eg 'A'
     * @return int Column index (A = 1)
     * @throws ColumnException
     */
    public function getColIndex(string $colStr): int
    {
        /** @var array */
        $indexCache = [];

        $charCount = strlen($colStr);
        if($charCount < 1 || $charCount > 11) {
            throw new ColumnException('Column row index cannot be less than 1 and longer than 10 characters.');
        }

        $colStr = strrev($colStr);

        for($i = 0; $i < $charCount; $i++) {

            if(isset($colStr[$i]) && array_key_exists($colStr[$i], $this->sourse)) {

                $index = $this->sourse[$colStr[$i]] * pow(26, $i);

                $key = strrev($colStr);

                if(!isset($indexCache[$key])) {
                    $indexCache = [$key => $index];
                } else {
                    $indexCache[$key] += $index;
                }

                continue;
            }

            throw new ColumnException('Index is invalid');
        }

        return $indexCache[strrev($colStr)];
    }
}
