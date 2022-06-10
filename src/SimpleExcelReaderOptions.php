<?php
declare(strict_types=1);

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Helper\EncodingHelper;
use OpenSpout\Reader\CSV\Options as CsvOptions;
use OpenSpout\Reader\ODS\Options as OdsOptions;
use OpenSpout\Reader\XLSX\Options as XlsxOptions;

trait SimpleExcelReaderOptions
{
    // CSV, ODS & Excel Options
    public bool $should_preserve_empty_rows = false;

    // CSV Options
    public string $field_delimiter = ',';
    public string $field_enclosure = '"';
    public string $encoding = EncodingHelper::ENCODING_UTF8;

    // ODS & Excel Options
    public bool $should_format_dates = false;

    // Excel Option
    public bool $should_use_1904_dates = false;

    public function use1904Dates(): self
    {
        $this->should_use_1904_dates = \true;

        return $this;
    }

    public function formatDates(): self
    {
        $this->should_format_dates = \true;

        return $this;
    }

    public function preserveEmptyRows(): self
    {
        $this->should_preserve_empty_rows = \true;

        return $this;
    }

    public function useEncoding(string $encoding): self
    {
        $this->encoding = $encoding;

        return $this;
    }

    public function useEnclosure(string $enclosure): self
    {
        $this->field_enclosure = $enclosure;

        return $this;
    }

    public function useDelimiter(string $delimiter): self
    {
        $this->field_delimiter = $delimiter;

        return $this;
    }

    public function getCsvOptions(): CsvOptions
    {
        $options = new CsvOptions();
        $options->SHOULD_PRESERVE_EMPTY_ROWS = $this->should_preserve_empty_rows;
        $options->FIELD_DELIMITER = $this->field_delimiter;
        $options->FIELD_ENCLOSURE = $this->field_enclosure;
        $options->ENCODING = $this->encoding;

        return $options;
    }

    public function getXlsxOptions(): XlsxOptions
    {
        $options = new XlsxOptions();
        $options->SHOULD_PRESERVE_EMPTY_ROWS = $this->should_preserve_empty_rows;
        $options->SHOULD_USE_1904_DATES = $this->should_use_1904_dates;
        $options->SHOULD_FORMAT_DATES = $this->should_format_dates;

        return $options;
    }

    public function getOdsOptions(): OdsOptions
    {
        $options = new OdsOptions();
        $options->SHOULD_PRESERVE_EMPTY_ROWS = $this->should_preserve_empty_rows;
        $options->SHOULD_FORMAT_DATES = $this->should_format_dates;

        return $options;
    }
}
