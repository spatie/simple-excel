<?php

namespace Spatie\SimpleExcel;

use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\CSV\Options as CSVOptions;
use OpenSpout\Writer\CSV\Writer;
use OpenSpout\Writer\WriterInterface;

class SimpleExcelWriter
{
    private WriterInterface $writer;

    private bool $processHeader = true;

    private bool $processingFirstRow = true;

    private int $numberOfRows = 0;

    private ?Style $headerStyle = null;

    protected CSVOptions $csvOptions;

    public static function create(
        string $file,
        string $type = '',
        callable $configureWriter = null,
        ?string $delimiter = null,
        ?bool $shouldAddBom = null,
    ): static {
        $simpleExcelWriter = new static(
            path: $file,
            type: $type,
            delimiter: $delimiter,
            shouldAddBom: $shouldAddBom,
        );

        $writer = $simpleExcelWriter->getWriter();

        if ($configureWriter) {
            $configureWriter($writer);
        }

        $writer->openToFile($file);

        return $simpleExcelWriter;
    }

    public static function createWithoutBom(string $file, string $type = ''): static
    {
        return static::create(
            file: $file,
            type: $type,
            shouldAddBom: false,
        );
    }

    public static function streamDownload(
        string $downloadName,
        string $type = '',
        callable $writerCallback = null,
        ?string $delimiter = null,
        ?bool $shouldAddBom = null,
    ): static {
        $simpleExcelWriter = new static($downloadName, $type, $delimiter, $shouldAddBom);

        $writer = $simpleExcelWriter->getWriter();

        if ($writerCallback) {
            $writerCallback($writer);
        }

        $writer->openToBrowser($downloadName);

        return $simpleExcelWriter;
    }

    protected function __construct(
        private string $path,
        string $type = '',
        ?string $delimiter = null,
        ?bool $shouldAddBom = null,
    ) {
        $this->csvOptions = new CSVOptions();

        $this->initWriter($path, $type);

        $this->addOptionsToWriter($path, $type, $delimiter, $shouldAddBom);
    }

    protected function initWriter(string $path, string $type, ?CSVOptions $options = null): void
    {
        $this->writer = empty($type) ?
            WriterFactory::createFromFile($path, $options) :
            WriterFactory::createFromType($type, $options);
    }

    protected function addOptionsToWriter(
        string $path,
        string $type = '',
        ?string $delimiter = null,
        ?bool $shouldAddBom = null,
    ): void {
        if (! $delimiter && $shouldAddBom) {
            return;
        }

        if (! $this->writer instanceof Writer) {
            return;
        }

        if ($delimiter !== null) {
            $this->csvOptions->FIELD_DELIMITER = $delimiter;
        }

        if ($shouldAddBom !== null) {
            $this->csvOptions->SHOULD_ADD_BOM = $shouldAddBom;
        }

        $this->initWriter($path, $type, $this->csvOptions);
    }

    public function getPath(): string
    {
        return $this->path;
    }

    public function getWriter(): WriterInterface
    {
        return $this->writer;
    }

    public function getNumberOfRows(): int
    {
        return $this->numberOfRows;
    }

    public function noHeaderRow(): static
    {
        $this->processHeader = false;

        return $this;
    }

    public function setHeaderStyle(Style $style): static
    {
        $this->headerStyle = $style;

        return $this;
    }

    public function addRow(Row|array $row, Style $style = null): static
    {
        if (is_array($row)) {
            if ($this->processHeader && $this->processingFirstRow) {
                $this->writeHeaderFromRow($row);
            }

            $row = Row::fromValues($row, $style);
        }

        $this->writer->addRow($row);
        $this->numberOfRows++;

        $this->processingFirstRow = false;

        return $this;
    }

    public function addRows(iterable $rows, Style $style = null): static
    {
        foreach ($rows as $row) {
            $this->addRow($row, $style);
        }

        return $this;
    }

    public function addHeader(array $header): self
    {
        $headerRow = Row::fromValues($header, $this->headerStyle);

        $this->writer->addRow($headerRow);
        $this->numberOfRows++;

        $this->processingFirstRow = false;

        return $this;
    }

    protected function writeHeaderFromRow(array $row): void
    {
        $headerValues = array_keys($row);

        $headerRow = Row::fromValues($headerValues, $this->headerStyle);

        $this->writer->addRow($headerRow);
        $this->numberOfRows++;
    }

    /**
     * Add a new sheet to the workbook.
     *
     * @param  string|null  $name The name of the sheet. If null, the name will be "SheetX" where X is the sheet index.
     *
     * @return $this
     */
    public function addNewSheetAndMakeItCurrent(?string $name = null): self
    {
        $this->writer->addNewSheetAndMakeItCurrent();
        $this->processingFirstRow = true;
        if ($name) {
            $this->nameCurrentSheet($name);
        }

        return $this;
    }

    /**
     * Sets the name for the current sheet.
     */
    public function nameCurrentSheet(string $name): self
    {
        $this->writer->getCurrentSheet()->setName($name);

        return $this;
    }

    public function toBrowser()
    {
        $this->writer->close();

        exit;
    }

    public function close(): void
    {
        $this->writer->close();
    }

    public function __destruct()
    {
        $this->close();
    }
}
