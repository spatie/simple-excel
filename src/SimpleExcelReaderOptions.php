<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Helper\EncodingHelper;
use OpenSpout\Reader\CSV\Options as CsvOptions;
use OpenSpout\Reader\ODS\Options as OdsOptions;
use OpenSpout\Reader\XLSX\Options as XlsxOptions;

trait SimpleExcelReaderOptions
{
    // CSV, ODS & Excel Options
    public static bool $should_preserve_empty_rows = false;

    // CSV Options
    public static string $field_delimiter = ',';
    public static string $field_enclosure = '"';
    public static string $encoding = EncodingHelper::ENCODING_UTF8;

    // ODS & Excel Options
    public static bool $should_format_dates = false;

    // Excel Option
    public static bool $should_use_1904_dates = false;

    public static function use1904Dates(): self
    {
        self::$should_use_1904_dates = \true;

        return self::getInstance();
    }

    public static function formatDates(): self
    {
        self::$should_use_1904_dates = \true;

        return self::getInstance();
    }

    public static function preserveEmptyRows(): self
    {
        self::$should_preserve_empty_rows = \true;

        return self::getInstance();
    }

    public static function useEncoding(string $encoding): self
    {
        self::$encoding = $encoding;

        return self::getInstance();
    }

    public static function useEnclosure(string $enclosure): self
    {
        self::$field_enclosure = $enclosure;

        return self::getInstance();
    }
    public static function useDelimiter(string $delimiter): self
    {
        self::$field_delimiter = $delimiter;

        return self::getInstance();
    }

    public function getCsvOptions(): CsvOptions
    {
        $options = new CsvOptions();
        $options->SHOULD_PRESERVE_EMPTY_ROWS = self::$should_preserve_empty_rows;
        $options->FIELD_DELIMITER = self::$field_delimiter;
        $options->FIELD_ENCLOSURE = self::$field_enclosure;
        $options->ENCODING = self::$encoding;

        // Need to reset the static variables to defaults for next instance
        $reset = new CsvOptions();
        self::$should_preserve_empty_rows = $reset->SHOULD_PRESERVE_EMPTY_ROWS;
        self::$field_delimiter = $reset->FIELD_DELIMITER;
        self::$field_enclosure = $reset->FIELD_ENCLOSURE;
        self::$encoding = $reset->ENCODING;

        return $options;
    }

    public function getXlsxOptions(): XlsxOptions
    {
        $options = new XlsxOptions();
        $options->SHOULD_PRESERVE_EMPTY_ROWS = self::$should_preserve_empty_rows;
        $options->SHOULD_USE_1904_DATES = self::$should_use_1904_dates;
        $options->SHOULD_FORMAT_DATES = self::$should_format_dates;

        // Need to reset the static variables to defaults for next instance
        $reset = new XlsxOptions();
        self::$should_preserve_empty_rows = $reset->SHOULD_PRESERVE_EMPTY_ROWS;
        self::$should_use_1904_dates = $reset->SHOULD_USE_1904_DATES;
        self::$should_format_dates = $reset->SHOULD_FORMAT_DATES;

        return $options;
    }

    public function getOdsOptions(): OdsOptions
    {
        $options = new OdsOptions();
        $options->SHOULD_PRESERVE_EMPTY_ROWS = self::$should_preserve_empty_rows;
        $options->SHOULD_FORMAT_DATES = self::$should_format_dates;

        return $options;
    }
}
