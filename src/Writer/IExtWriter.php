<?php

namespace arsenweb\Writer;

use arsenweb\Reader\IExtReader;

interface IExtWriter
{
    public static function setReader(IExtReader $reader): IExtWriter;

    public function setString(string $mask, string $value): IExtWriter;

    public function setInt(string $mask, int $value): IExtWriter;

    public function setFloat(string $mask, $value, int $precision = 2): IExtWriter;

    public function setArray(
        string $mask,
        array $items,
        ?callable $fn = null,
        string $direction = self::DIRECTION_DOWN
    ): IExtWriter;

    public function setAssociativeArray(string $mask, array $rows, ?callable $fn = null, ?callable $itemHandler = null): IExtWriter;

    public function save(string $pathToFile, ?string $fileType = null): bool;

    public function getContent(string $fileFileType): string;

    public function getBase64(string $fileType): string;
}
