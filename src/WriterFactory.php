<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Writer\CSV\{Options as CSVOptions, Writer as CSVWriter};
use OpenSpout\Writer\ODS\{Options as ODSOptions, Writer as ODSWriter};
use OpenSpout\Writer\WriterInterface;
use OpenSpout\Writer\XLSX\{Options as XLSXOptions, Writer as XLSXWriter};

/**
 * @internal overwritten from openspout/openspout so we can pass Options to the Writer classes
 * Original: \OpenSpout\Writer\Common\Creator\ReaderFactory
 */
class WriterFactory
{
    use SpreadsheetFactoryTrait;

    private const FILE_EXTENSION_MAP = [
        'csv' => CSVWriter::class,
        'xlsx' => XLSXWriter::class,
        'ods' => ODSWriter::class,
    ];

    /**
     * This creates an instance of the appropriate writer, given the extension of the file to be written.
     *
     * @param string $path The path to the spreadsheet file. Supported extensions are .csv,.ods and .xlsx
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     */
    public static function createFromFile(
        string $path,
        CSVOptions|XLSXOptions|ODSOptions|null $options = null
    ): WriterInterface {
        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));

        return self::resolveFromType(
            $extension,
            self::FILE_EXTENSION_MAP,
            $options,
            "No writers supporting the given type: {$extension}"
        );
    }

    /**
     * @deprecated use createFromFile() instead
     *
     * @throws UnsupportedTypeException
     */
    public static function createFromType(
        string $writerType,
        CSVOptions|XLSXOptions|ODSOptions|null $options = null
    ): WriterInterface {
        return self::resolveFromType(
            strtolower($writerType),
            self::FILE_EXTENSION_MAP,
            $options,
            "No writers supporting the given type: {$writerType}"
        );
    }
}
