<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Exception\IOException;
use OpenSpout\Common\Exception\UnsupportedTypeException;
use OpenSpout\Reader\CSV\Options as CSVOptions;
use OpenSpout\Reader\CSV\Reader as CSVReader;
use OpenSpout\Reader\ODS\Options as ODSOptions;
use OpenSpout\Reader\ODS\Reader as ODSReader;
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Reader\XLSX\Options as XLSXOptions;
use OpenSpout\Reader\XLSX\Reader as XLSXReader;

/**
 * @internal overwritten from openspout/openspout so we can pass Options to the Reader classes
 * Original: \OpenSpout\Reader\Common\Creator\ReaderFactory
 */
class ReaderFactory
{
    /**
     * Creates a reader by file extension.
     *
     * @param string $path The path to the spreadsheet file. Supported extensions are .csv,.ods and .xlsx
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     */
    public static function createFromFile(
        string $path,
        CSVOptions|XLSXOptions|ODSOptions|null $options = null
    ): ReaderInterface {
        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));

        return match ($extension) {
            'csv' => new CSVReader($options),
            'xlsx' => new XLSXReader($options),
            'ods' => new ODSReader($options),
            default => throw new UnsupportedTypeException('No readers supporting the given type: '.$extension),
        };
    }

    /**
     * Creates a reader by mime type.
     *
     * @param string $path the path to the spreadsheet file
     *
     * @throws \OpenSpout\Common\Exception\UnsupportedTypeException
     * @throws \OpenSpout\Common\Exception\IOException
     */
    public static function createFromFileByMimeType(
        string $path,
        CSVOptions|XLSXOptions|ODSOptions|null $options = null
    ): ReaderInterface {
        if (! file_exists($path)) {
            throw new IOException("Could not open {$path} for reading! File does not exist.");
        }

        $mime_type = mime_content_type($path);

        return match ($mime_type) {
            'application/csv', 'text/csv', 'text/plain' => new CSVReader($options),
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' => new XLSXReader($options),
            'application/vnd.oasis.opendocument.spreadsheet' => new ODSReader($options),
            default => throw new UnsupportedTypeException('No readers supporting the given type: '.$mime_type),
        };
    }

    /**
     * @deprecated use createFromFileByMimeType() or createFromFile() instead
     *
     * @throws UnsupportedTypeException
     */
    public static function createFromType(
        string $readerType,
        CSVOptions|XLSXOptions|ODSOptions|null $options = null
    ): ReaderInterface {
        return match ($readerType) {
            'csv' => new CSVReader($options),
            'xlsx' => new XLSXReader($options),
            'ods' => new ODSReader($options),
            default => throw new UnsupportedTypeException('No readers supporting the given type: ' . $readerType),
        };
    }
}
