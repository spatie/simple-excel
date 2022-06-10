<?php
declare(strict_types=1);

namespace Spatie\SimpleExcel;

use Illuminate\Support\LazyCollection;
use InvalidArgumentException;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Exception\UnsupportedTypeException;
use OpenSpout\Reader\CSV\Reader as CsvReader;
use OpenSpout\Reader\ODS\Reader as OdsReader;
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Reader\RowIteratorInterface;
use OpenSpout\Reader\SheetInterface;
use OpenSpout\Reader\XLSX\Reader as XlsxReader;

class SimpleExcelReader
{
    use SimpleExcelReaderOptions;

    protected string $path;

    protected ReaderInterface $reader;

    protected RowIteratorInterface $rowIterator;

    protected int $sheetNumber = 1;

    protected bool $processHeader = true;

    protected bool $trimHeader = false;

    protected bool $headersToSnakeCase = false;

    protected ?string $trimHeaderCharacters = null;

    protected mixed $formatHeadersUsing = null;

    protected ?array $headers = null;

    protected int $headerOnRow = 0;

    protected int $skip = 0;

    protected int $limit = 0;

    protected bool $useLimit = false;

    public static function options(): self
    {
        return new static();
    }

    public static function create(string $file, ?string $type = \null)
    {
        $extension = strtolower(pathinfo($file, PATHINFO_EXTENSION));

        if ($type !== \null) {
            $extension = strtolower($type);
        }

        return match ($extension) {
            'csv' => (new self)->openCsv($file),
            'ods' => (new self)->openOds($file),
            'xlsx' => (new self)->openXlsx($file),
            default => throw new UnsupportedTypeException('No readers supporting the given type: '.$extension),
        };
    }

    public static function createCsv(string $file)
    {
        return (new self)->openCsv($file);
    }

    public function openCsv(string $file)
    {
        $this->path = $file;

        $this->reader = new CsvReader($this->getCsvOptions());

        return $this;
    }

    public static function createOds(string $file)
    {
        return (new self)->openOds($file);
    }

    public function openOds(string $file)
    {
        $this->path = $file;

        $this->reader = new OdsReader($this->getOdsOptions());

        return $this;
    }

    public static function createXlsx(string $file)
    {
        return (new self)->openXlsx($file);
    }

    public function openXlsx(string $file)
    {
        $this->path = $file;

        $this->reader = new XlsxReader($this->getXlsxOptions());

        return $this;
    }

    public function getPath(): string
    {
        return $this->path;
    }

    public function headerOnRow(int $headerRow): self
    {
        $this->headerOnRow = $headerRow;

        return $this;
    }

    public function noHeaderRow(): self
    {
        $this->processHeader = false;

        return $this;
    }

    public function trimHeaderRow(string $characters = null): self
    {
        $this->trimHeader = true;
        $this->trimHeaderCharacters = $characters;

        return $this;
    }

    public function formatHeadersUsing(callable $callback): self
    {
        $this->formatHeadersUsing = $callback;

        return $this;
    }

    public function headersToSnakeCase(): self
    {
        $this->headersToSnakeCase = true;

        return $this;
    }

    public function getReader(): ReaderInterface
    {
        return $this->reader;
    }

    public function skip(int $count): SimpleExcelReader
    {
        $this->skip = $count;

        return $this;
    }

    public function take(int $count): SimpleExcelReader
    {
        $this->limit = $count;
        $this->useLimit = true;

        return $this;
    }

    public function fromSheet(int $sheetNumber): SimpleExcelReader
    {
        $this->sheetNumber = $sheetNumber;
        $this->headers = null;

        return $this;
    }

    public function getRows(): LazyCollection
    {
        $sheet = $this->getSheet();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        $this->getHeaders();

        if ($this->processHeader) {
            $this->rowIterator->next();
        }

        return LazyCollection::make(function () {
            while ($this->rowIterator->valid() && $this->skip && $this->skip--) {
                $this->rowIterator->next();
            }
            while ($this->rowIterator->valid() && (! $this->useLimit || $this->limit--)) {
                $row = $this->rowIterator->current();

                yield $this->getValueFromRow($row);

                $this->rowIterator->next();
            }
        });
    }

    public function getHeaders(): ?array
    {
        if (! $this->processHeader) {
            return null;
        }

        if ($this->headers) {
            return $this->headers;
        }

        $sheet = $this->getSheet();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        /** @var \OpenSpout\Common\Entity\Row|null $headerRow */
        $headerRow = $this->rowIterator->current();

        if ($this->headerOnRow > 0) {
            $skip = $this->headerOnRow;
            while ($skip--) {
                $this->rowIterator->next();
            }
            $headerRow = $this->rowIterator->current();
        }

        if (is_null($headerRow)) {
            $this->noHeaderRow();

            return null;
        }

        $this->headers = $this->processHeaderRow($headerRow->toArray());

        return $this->headers;
    }

    public function close()
    {
        $this->reader->close();
    }

    protected function processHeaderRow(array $headers): array
    {
        if ($this->trimHeader) {
            $headers = $this->convertHeaders([$this, 'trim'], $headers);
        }

        if ($this->headersToSnakeCase) {
            $headers = $this->convertHeaders([$this, 'toSnakecase'], $headers);
        }

        if ($this->formatHeadersUsing) {
            $headers = $this->convertHeaders($this->formatHeadersUsing, $headers);
        }

        return $headers;
    }

    protected function convertHeaders(callable $callback, array $headers): array
    {
        return array_map(function ($header) use ($callback) {
            return call_user_func($callback, $header);
        }, $headers);
    }

    public function headerRowFormatter(callable $callback)
    {
        $this->headerRowFormatter = $callback;

        return $this;
    }

    protected function trim(string $header): string
    {
        $arguments[] = $header;

        if (! is_null($this->trimHeaderCharacters)) {
            $arguments[] = $this->trimHeaderCharacters;
        }

        return call_user_func_array('trim', $arguments);
    }

    protected function toSnakeCase(string $header): string
    {
        return str_replace(
            ' ',
            '_',
            strtolower(preg_replace('/(?<=\d)(?=[A-Za-z])|(?<=[A-Za-z])(?=\d)|(?<=[a-z])(?=[A-Z])/', '_', trim($header)))
        );
    }

    /**
     * @param \OpenSpout\Common\Entity\Row $row
     *
     * @return array<int,string>
     */
    protected function getValueFromRow(Row $row): array
    {
        /** @var array<int,string> $values */
        $values = $row->toArray();
        ksort($values);

        if (! $this->processHeader) {
            return $values;
        }

        $values = array_slice($values, 0, count($this->headers));

        while (count($values) < count($this->headers)) {
            $values[] = '';
        }

        return array_combine($this->headers, $values);
    }

    protected function getSheet(): SheetInterface
    {
        $this->reader->open($this->path);

        foreach ($this->reader->getSheetIterator() as $key => $sheet) {
            if ($key === $this->sheetNumber) {
                break;
            }
        }

        if ($this->sheetNumber !== $key) {
            throw new InvalidArgumentException("Sheet {$key} does not exist in {$this->path}.");
        }

        return $sheet;
    }

    public function __destruct()
    {
        $this->close();
    }
}
