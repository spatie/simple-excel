# Upgrade guide

## Upgrading from 2.x to 3.0


### Most notable changes

1. Add support for openspout/openspout v4
2. Drop support for openspout/openspout v3
3. Add type hinting
4. Removed `useDelimiter` on SimpleExcelWriter
5. Removed `headerRowFormatter` on SimpleExcelReader

### Classes have been moved

- `\Box\Spout\Common\Entity\Row` should be replaced with `\OpenSpout\Common\Entity\Row`
- `\Box\Spout\Common\Entity\Style\Style` should be replaced with `OpenSpout\Common\Entity\Style\Style`

### Removed `useDelimiter()` on SimpleExcelWriter

In v3 there was a method to set a delimiter. Now you should pass this as parameter to the constructor.

Change
```php
$reader = SimpleExcelWriter::create($file)->useDelimiter(';');
```

To
```php
 $writer = SimpleExcelWriter::create(file: $file, delimiter: ';');
```

### Replace StyleBuilder with Style

In OpenSpout v4 the `StyleBuilder` is removed and integrated inside the `Style` class.

Update code like this...

```php
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\Common\Creator\Style\StyleBuilder;

$builder = new StyleBuilder();
$builder
    ->setFontBold()
    ->setFontName('Sans');
```

... to ...

```php
use OpenSpout\Common\Entity\Style\Style;

$style = new Style();
$style
    ->setFontBold()
    ->setFontName('Sans');
```

### Deprecated setting the type manually

In v4 of openspout/openspout it is no longer possible to explicitly set the type.
We still have support for this, but we'll deprecate the method.

```php
$reader = SimpleExcelReader::create('php://input', 'csv');
```

```php
 $writer = SimpleExcelWriter::create('php://output', 'csv');
```
