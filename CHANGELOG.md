# Changelog

All notable changes to `simple-excel` will be documented in this file

## 2.4.0 - 2022-11-09

### What's Changed

- Refactor tests to Pest by @alexmanase in https://github.com/spatie/simple-excel/pull/123
- Add optional styling when adding multiple rows to xslx by @chrisdicarlo in https://github.com/spatie/simple-excel/pull/122
- Add new sheet to XLSX with Header-Row by @red-freak in https://github.com/spatie/simple-excel/pull/124

### New Contributors

- @alexmanase made their first contribution in https://github.com/spatie/simple-excel/pull/123
- @chrisdicarlo made their first contribution in https://github.com/spatie/simple-excel/pull/122
- @red-freak made their first contribution in https://github.com/spatie/simple-excel/pull/124

**Full Changelog**: https://github.com/spatie/simple-excel/compare/2.3.0...2.4.0

## 2.3.0 - 2022-10-11

### What's Changed

- Use field delimiter only for csv reader by @dehbka in https://github.com/spatie/simple-excel/pull/111
- Add function fromSheetName to read on specific spreadsheet name by @SpyBott in https://github.com/spatie/simple-excel/pull/117

### New Contributors

- @dehbka made their first contribution in https://github.com/spatie/simple-excel/pull/111
- @SpyBott made their first contribution in https://github.com/spatie/simple-excel/pull/117

**Full Changelog**: https://github.com/spatie/simple-excel/compare/2.2.0...2.3.0

## 2.2.0 - 2022-09-02

### What's Changed

- Documentation for fromSheet and headerOnRow by @dakira in https://github.com/spatie/simple-excel/pull/102
- Ability to set custom headers (whether or not headers exist in the file) by @kitbs in https://github.com/spatie/simple-excel/pull/103

### New Contributors

- @kitbs made their first contribution in https://github.com/spatie/simple-excel/pull/103

**Full Changelog**: https://github.com/spatie/simple-excel/compare/2.1.0...2.2.0

## 2.1.0 - 2022-06-09

### What's Changed

- Implement headerOnRow method by @ArondeParon in https://github.com/spatie/simple-excel/pull/89

### New Contributors

- @ArondeParon made their first contribution in https://github.com/spatie/simple-excel/pull/89

**Full Changelog**: https://github.com/spatie/simple-excel/compare/2.0.0...2.1.0

## 2.0.0 - 2022-06-06

- move from box/spout to openspout/openspout

## 1.15.1 - 2022-05-11

## What's Changed

- Update README.md by @vdvcoder in https://github.com/spatie/simple-excel/pull/80
- Typo by @saurabhsharma2u in https://github.com/spatie/simple-excel/pull/83
- Allow for selecting sheet by number by @dakira in https://github.com/spatie/simple-excel/pull/86

## New Contributors

- @vdvcoder made their first contribution in https://github.com/spatie/simple-excel/pull/80
- @saurabhsharma2u made their first contribution in https://github.com/spatie/simple-excel/pull/83
- @dakira made their first contribution in https://github.com/spatie/simple-excel/pull/86

**Full Changelog**: https://github.com/spatie/simple-excel/compare/1.15.0...1.15.1

## 1.15.0 - 2022-01-12

- support Laravel 9

## 1.14.1 - 2021-06-11

- Allow trimming empty header titles (#64)

## 1.14.0 - 2021-06-10

- Allow all forms of callables to format header (#63)
- drop support for PHP 7

## 1.13.1 - 2021-03-26

- make sure `getHeaders()` take delimiter in account (#57)

## 1.13.0 - 2020-01-18

- add `getHeaders()` (#52)

## 1.12.0 - 2020-12-30

- add `headersToSnakeCase` and `formatHeadersUsing`

## 1.11.0 - 2020-12-29

- enable disabling BOM on writer (#48)

## 1.10.2 - 2020-12-28

- use setHeaderStyle fluently (#47)

## 1.10.1 - 2020-12-27

- enable header row trimming (#46)

## 1.10.0 - 2020-12-08

- allow setting the writer/reader type manually (#43)

## 1.9.1 - 2020-11-30

- add support for PHP 8

## 1.9.0 - 2020-10-30

- add Header Styling Method (#39)

## 1.8.1 - 2020-10-08

- fix `skip` method

## 1.8.0 - 2020-10-04

- add 'take' and 'skip' functions to reader (#35)

## 1.7.1 - 2020-09-08

- allow Laravel 8

## 1.7.0 - 2020-08-19

- make `addRows` chainable

## 1.6.0 - 2020-07-16

- add `addRows`

## 1.5.0 - 2020-07-01

- wrong tag, please ignore

## 1.4.0 - 2020-04-15

- Add `useDelimiter` method for `SimpleExcelWriter` (#25)

## 1.3.1 - 2019-02-17

- Fix columns being returned in the wrong order

## 1.3.0 - 2019-01-02

- drop support for PHP 7.3

## 1.2.2 - 2019-11-29

- make sure `streamDownload` does not create a file

## 1.2.0 - 2019-11-29

- add `streamDownload` and `toBrowser`

## 1.1.0 - 2019-10-27

- add `getPath`

## 1.0.0 - 2019-10-26

- initial release
