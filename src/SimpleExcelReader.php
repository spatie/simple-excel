<?php
namespace Spatie\SimpleExcel;

use Illuminate\Support\LazyCollection;
use InvalidArgumentException;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Reader\Common\Creator\ReaderEntityFactory;
use OpenSpout\Reader\Common\Creator\ReaderFactory;
use OpenSpout\Reader\CSV\Reader as CSVReader;
use OpenSpout\Reader\IteratorInterface;
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Reader\SheetInterface;

class SimpleExcelReader
{
    protected string $path;
    protected string $type;
    protected ReaderInterface $reader;
    protected IteratorInterface $rowIterator;
    protected int $sheetNumber = 1;
    protected string $sheetName = "";
    protected bool $searchSheetByName = false;
    protected bool $processHeader = true;
    protected bool $trimHeader = false;
    protected bool $headersToSnakeCase = false;
    protected ?string $trimHeaderCharacters = null;
    protected mixed $formatHeadersUsing = null;
    protected ?array $headers = null;
    protected int $headerOnRow = 0;
    protected ?array $customHeaders = [];
    protected int $skip = 0;
    protected int $limit = 0;
    protected bool $useLimit = false;

    public static function create(string $file, string $type = '')
    {
        return new static($file, $type);
    }

    public function __construct(string $path, string $type = '')
    {
        $this->path = $path;

        $this->type = $type;

        $this->reader = $type ?
            ReaderFactory::createFromType($type) :
            ReaderEntityFactory::createReaderFromFile($this->path);
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

    public function useHeaders(array $headers): self
    {
        $this->customHeaders = $headers;

        return $this;
    }

    public function useDelimiter(string $delimiter): self
    {
        if ($this->reader instanceof CSVReader) {
            $this->reader->setFieldDelimiter($delimiter);
        }

        return $this;
    }

    public function useFieldEnclosure(string $fieldEnclosure): self
    {
        $this->reader->setFieldEnclosure($fieldEnclosure);

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
        $this->searchSheetByName = false;

        return $this;
    }

    public function fromSheetName(string $sheetName): SimpleExcelReader
    {
        $this->searchSheetByName = true;
        $this->sheetName = $sheetName;
        $this->headers = null;

        return $this;
    }

    public function getRows(): LazyCollection
    {
        $sheet = $this->getSheet();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

        if ($this->processHeader) {
            $this->getHeaders();
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
            if ($this->customHeaders) {
                return $this->customHeaders;
            }

            return null;
        }

        if ($this->headers) {
            return $this->headers;
        }

        $sheet = $this->getSheet();

        $this->rowIterator = $sheet->getRowIterator();

        $this->rowIterator->rewind();

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

        if ($this->customHeaders) {
            return $this->customHeaders;
        }

        return $this->headers;
    }

    public function getOriginalHeaders(): ?array
    {
        $this->getHeaders();

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

    protected function getValueFromRow(Row $row): array
    {
        $values = $row->toArray();
        ksort($values);

        $headers = $this->customHeaders ?: $this->headers;

        if (! $headers) {
            return $values;
        }

        $values = array_slice($values, 0, count($headers));

        while (count($values) < count($headers)) {
            $values[] = '';
        }

        return array_combine($headers, $values);
    }

    protected function getSheet(): SheetInterface
    {
        $this->reader->open($this->path);
        $sheet = ($this->searchSheetByName) ? $this->getActiveSheetByName() : $this->getActiveSheetByIndex();

        return $sheet;
    }

    protected function getActiveSheetByName(): SheetInterface
    {
        foreach ($this->reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheetName != "" && $this->sheetName === $sheet->getName()) {
                break;
            }
        }
        if ($this->sheetName != "" && $this->sheetName !== $sheet->getName()) {
            throw new InvalidArgumentException("Sheet name {$this->sheetName} does not exist in {$this->path}.");
        }

        return $sheet;
    }

    protected function getActiveSheetByIndex(): SheetInterface
    {
        foreach ($this->reader->getSheetIterator() as $key => $sheet) {
            if ($key === $this->sheetNumber) {
                break;
            }
        }
        if ($this->sheetNumber !== $key) {
            throw new InvalidArgumentException("Sheet Index {$key} does not exist in {$this->path}.");
        }

        return $sheet;
    }

    public function __destruct()
    {
        $this->close();
    }
}
