<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Exception\IOException;
use OpenSpout\Reader\CSV\{Options as CSVOptions, Reader as CSVReader};
use OpenSpout\Reader\ODS\{Options as ODSOptions, Reader as ODSReader};
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Reader\XLSX\{Options as XLSXOptions, Reader as XLSXReader};

/**
 * @internal overwritten from openspout/openspout so we can pass Options to the Reader classes
 * Original: \OpenSpout\Reader\Common\Creator\ReaderFactory
 */
class ReaderFactory
{
    use SpreadsheetFactoryTrait;

    private const FILE_EXTENSION_MAP = [
        'csv' => CSVReader::class,
        'xlsx' => XLSXReader::class,
        'ods' => ODSReader::class,
    ];

    private const MIME_TYPE_MAP = [
        'application/csv' => CSVReader::class,
        'text/csv' => CSVReader::class,
        'text/plain' => CSVReader::class,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' => XLSXReader::class,
        'application/vnd.oasis.opendocument.spreadsheet' => ODSReader::class,
    ];

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

        return self::resolveFromType(
            $extension,
            self::FILE_EXTENSION_MAP,
            $options,
            "No readers supporting the given type: {$extension}"
        );
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

        $mimeType = mime_content_type($path);
        if ($mimeType === false) {
            throw new IOException("Could not determine mime type for {$path}");
        }

        return self::resolveFromType(
            $mimeType,
            self::MIME_TYPE_MAP,
            $options,
            "No readers supporting the given type: {$mimeType}"
        );
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
        return self::resolveFromType(
            strtolower($readerType),
            self::FILE_EXTENSION_MAP,
            $options,
            "No readers supporting the given type: {$readerType}"
        );
    }
}
