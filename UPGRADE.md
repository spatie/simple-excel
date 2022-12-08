# Upgrade guide

## Upgrading from 3.x to 4.0


### Most notable changes

1. Add support for openspout/openspout v4
2. Drop support for openspout/openspout v3

### Classes have been moved

- `\Box\Spout\Common\Entity\Row` should be replaced with `\OpenSpout\Common\Entity\Row`
- `\Box\Spout\Common\Entity\Style\Style` should be replaced with `OpenSpout\Common\Entity\Style\Style`

### Remove `useDelimiter()`

In v3 there was a method to set a delimiter. Now you should pass this as parameter.

Change
```php
$reader = SimpleExcelWriter::create($file)->useDelimiter(';');
```

To
```php
 $writer = SimpleExcelWriter::create(file: $file, delimiter: ';');
```

### Drop support for setting the type manually

In v3 of this package it was possible to explicitly set the type. From now in this is not possible anymore

Change
```php
$reader = SimpleExcelReader::create('php://input', 'csv');
```

To
```php
 $writer = SimpleExcelWriter::create('php://output', 'csv');
```
