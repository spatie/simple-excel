<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Exception\UnsupportedTypeException;
use OpenSpout\Writer\CSV\Options as CSVOptions;
use OpenSpout\Writer\CSV\Writer as CSVWriter;
use OpenSpout\Writer\ODS\Options as ODSOptions;
use OpenSpout\Writer\ODS\Writer as ODSWriter;
use OpenSpout\Writer\WriterInterface;
use OpenSpout\Writer\XLSX\Options as XLSXOptions;
use OpenSpout\Writer\XLSX\Writer as XLSXWriter;

/**
 * @internal overwritten from openspout/openspout so we can pass Options to the Writer classes
 * Original: \OpenSpout\Writer\Common\Creator\ReaderFactory
 */
class WriterFactory
{
    /**
     * This creates an instance of the appropriate writer, given the extension of the file to be written.
     *
     * @param string $path The path to the spreadsheet file. Supported extensions are .csv,.ods and .xlsx
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     */
    public static function createFromFile(
        string $path,
        CSVOptions|XLSXOptions|ODSOptions|null $options = null,
    ): WriterInterface {
        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));

        return match ($extension) {
            'csv' => new CSVWriter($options),
            'xlsx' => new XLSXWriter($options),
            'ods' => new ODSWriter($options),
            default => throw new UnsupportedTypeException('No writers supporting the given type: '.$extension),
        };
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
        return match ($writerType) {
            'csv' => new CSVWriter($options),
            'xlsx' => new XLSXWriter($options),
            'ods' => new ODSWriter($options),
            default => throw new UnsupportedTypeException('No writers supporting the given type: ' . $writerType),
        };
    }
}
