<?php

namespace MewesK\TwigSpreadsheetBundle\Tests\Twig;

use PhpOffice\PhpSpreadsheet\Cell\DataType;

class OdsXlsXlsxTwigTest extends BaseTwigTest
{
    public function formatProvider(): array
    {
        return [['ods'], ['xls'], ['xlsx']];
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testBlock(string $format): void
    {
        $document = $this->getDocument('block', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testBlockOverrideCell(string $format): void
    {
        $document = $this->getDocument('blockOverrideCell', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('Bar2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testBlockOverrideContent(string $format): void
    {
        $document = $this->getDocument('blockOverrideContent', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Foo2', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testBlockOverrideRow(string $format): void
    {
        $document = $this->getDocument('blockOverrideRow', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello2', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World2', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testBlockOverrideSheet(string $format): void
    {
        $document = $this->getDocument('blockOverrideSheet', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test2');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello3', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World3', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertNotSame('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertNotSame('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testCellFormula(string $format): void
    {
        $document = $this->getDocument('cellFormula', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('=A1*B1+2', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertTrue($sheet->getCell('A2')->isFormula(), 'Unexpected value in isFormula');
        static::assertSame(1337, (int) $sheet->getCell('A2')->getCalculatedValue(), 'Unexpected calculated value in A2');

        static::assertSame('=SUM(A1:B1)', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
        static::assertTrue($sheet->getCell('A3')->isFormula(), 'Unexpected value in isFormula');
        static::assertSame(669.5, $sheet->getCell('A3')->getCalculatedValue(), 'Unexpected calculated value in A3');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testCellIndex(string $format): void
    {
        $document = $this->getDocument('cellIndex', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('Hello', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertNotSame('Bar', $sheet->getCell('C1')->getValue(), 'Unexpected value in C1');
        static::assertSame('Lorem', $sheet->getCell('C1')->getValue(), 'Unexpected value in C1');
        static::assertSame('Ipsum', $sheet->getCell('D1')->getValue(), 'Unexpected value in D1');
        static::assertSame('World', $sheet->getCell('E1')->getValue(), 'Unexpected value in E1');

        static::assertSame('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A1');
        static::assertSame('Bar', $sheet->getCell('B2')->getValue(), 'Unexpected value in B1');
        static::assertSame('Lorem', $sheet->getCell('C2')->getValue(), 'Unexpected value in C1');
        static::assertSame('Ipsum', $sheet->getCell('D2')->getValue(), 'Unexpected value in D1');
        static::assertSame('Hello', $sheet->getCell('E2')->getValue(), 'Unexpected value in E1');
        static::assertSame('World', $sheet->getCell('F2')->getValue(), 'Unexpected value in F1');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testCellProperties(string $format): void
    {
        $document = $this->getDocument('cellProperties', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        $cell = $sheet->getCell('A1');
        static::assertNotNull($cell, 'Cell does not exist');

        $dataValidation = $cell->getDataValidation();
        static::assertNotNull($dataValidation, 'DataValidation does not exist');

        $style = $cell->getStyle();
        static::assertNotNull($style, 'Style does not exist');

        static::assertSame(42, $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('A2')->getDataType(), 'Unexpected value in dataType');

        static::assertSame('42', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
        static::assertSame(DataType::TYPE_STRING, $sheet->getCell('B2')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(42, $sheet->getCell('C2')->getValue(), 'Unexpected value in C2');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('C2')->getDataType(), 'Unexpected value in dataType');

        static::assertSame('007', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
        static::assertSame(DataType::TYPE_STRING, $sheet->getCell('A3')->getDataType(), 'Unexpected value in dataType');

        static::assertSame('007', $sheet->getCell('B3')->getValue(), 'Unexpected value in B3');
        static::assertSame(DataType::TYPE_STRING, $sheet->getCell('B3')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(7, $sheet->getCell('C3')->getValue(), 'Unexpected value in C3');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('C3')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(0.1337, $sheet->getCell('A4')->getValue(), 'Unexpected value in A4');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('A4')->getDataType(), 'Unexpected value in dataType');

        static::assertSame('0.13370', $sheet->getCell('B4')->getValue(), 'Unexpected value in B4');
        static::assertSame(DataType::TYPE_STRING, $sheet->getCell('B4')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(0.1337, $sheet->getCell('C4')->getValue(), 'Unexpected value in C4');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('C4')->getDataType(), 'Unexpected value in dataType');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testCellValue(string $format): void
    {
        $document = $this->getDocument('cellValue', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame(667.5, $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('A1')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(2, $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('B1')->getDataType(), 'Unexpected value in dataType');

        static::assertSame('007', $sheet->getCell('C1')->getValue(), 'Unexpected value in C1');
        static::assertSame(DataType::TYPE_STRING, $sheet->getCell('C1')->getDataType(), 'Unexpected value in dataType');

        static::assertSame('foo', $sheet->getCell('D1')->getValue(), 'Unexpected value in D1');
        static::assertSame(DataType::TYPE_STRING, $sheet->getCell('D1')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(0.42, $sheet->getCell('E1')->getValue(), 'Unexpected value in E1');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('E1')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(21, $sheet->getCell('F1')->getValue(), 'Unexpected value in F1');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('F1')->getDataType(), 'Unexpected value in dataType');

        static::assertSame(1.2, $sheet->getCell('G1')->getValue(), 'Unexpected value in G1');
        static::assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('G1')->getDataType(), 'Unexpected value in dataType');

        static::assertSame('BAR', $sheet->getCell('H1')->getValue(), 'Unexpected value in H1');
        static::assertSame(DataType::TYPE_STRING, $sheet->getCell('H1')->getDataType(), 'Unexpected value in dataType');
    }

    /**
     * The following attributes are not supported by the readers and therefore cannot be tested:
     * $security->getLockRevision() -> true
     * $security->getLockStructure() -> true
     * $security->getLockWindows() -> true
     * $security->getRevisionsPassword() -> 'test'
     * $security->getWorkbookPassword() -> 'test'.
     *
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentProperties(string $format): void
    {
        $document = $this->getDocument('documentProperties', $format);
        static::assertNotNull($document, 'Document does not exist');

        $properties = $document->getProperties();
        static::assertNotNull($properties, 'Properties do not exist');

        // +/- 24h range to allow possible timezone differences (946684800)
        static::assertGreaterThanOrEqual(946_598_400, $properties->getCreated(), 'Unexpected value in created');
        static::assertLessThanOrEqual(946_771_200, $properties->getCreated(), 'Unexpected value in created');
        static::assertSame('Test creator', $properties->getCreator(), 'Unexpected value in creator');

        $defaultStyle = $document->getDefaultStyle();
        static::assertNotNull($defaultStyle, 'DefaultStyle does not exist');

        static::assertSame('Test description', $properties->getDescription(), 'Unexpected value in description');
        // +/- 24h range to allow possible timezone differences (946684800)
        static::assertGreaterThanOrEqual(946_598_400, $properties->getModified(), 'Unexpected value in modified');
        static::assertLessThanOrEqual(946_771_200, $properties->getModified(), 'Unexpected value in modified');

        $security = $document->getSecurity();
        static::assertNotNull($security, 'Security does not exist');

        static::assertSame('Test subject', $properties->getSubject(), 'Unexpected value in subject');
        static::assertSame('Test title', $properties->getTitle(), 'Unexpected value in title');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentWhitespace(string $format): void
    {
        $document = $this->getDocument('documentWhitespace', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('Bar', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Hello', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('World', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testInclude(string $format): void
    {
        $document = $this->getDocument('include', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello1', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World1', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Hello2', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('World2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');

        $sheet = $document->getSheetByName('Test2');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello3', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World3', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testMacro(string $format): void
    {
        $document = $this->getDocument('macro', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello1', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World1', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Hello2', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('World2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');

        $sheet = $document->getSheetByName('Test2');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello3', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World3', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');

        $sheet = $document->getSheetByName('Test3');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello4', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World4', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testRowIndex(string $format): void
    {
        $document = $this->getDocument('rowIndex', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertNotSame('Bar', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
        static::assertSame('Lorem', $sheet->getCell('A3')->getValue(), 'Unexpected value in A3');
        static::assertSame('Ipsum', $sheet->getCell('A4')->getValue(), 'Unexpected value in A4');
        static::assertSame('Hello', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('World', $sheet->getCell('A5')->getValue(), 'Unexpected value in A5');

        static::assertSame('Foo', $sheet->getCell('A6')->getValue(), 'Unexpected value in A6');
        static::assertSame('Bar', $sheet->getCell('A7')->getValue(), 'Unexpected value in A7');
        static::assertSame('Lorem', $sheet->getCell('A8')->getValue(), 'Unexpected value in A8');
        static::assertSame('Ipsum', $sheet->getCell('A9')->getValue(), 'Unexpected value in A9');
        static::assertSame('Hello', $sheet->getCell('A10')->getValue(), 'Unexpected value in A10');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testSheet(string $format): void
    {
        $document = $this->getDocument('documentSimple', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');
        static::assertSame($sheet, $document->getActiveSheet(), 'Sheets are not equal');
        static::assertCount(1, $document->getAllSheets(), 'Unexpected sheet count');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testSheetComplex(string $format): void
    {
        $document = $this->getDocument('sheetComplex', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test 1');
        static::assertNotNull($sheet, 'Sheet "Test 1" does not exist');
        static::assertSame('Foo', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('Bar', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');

        $sheet = $document->getSheetByName('Test 2');
        static::assertNotNull($sheet, 'Sheet "Test 2" does not exist');
        static::assertSame('Hello World', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
    }

    /**
     * The following attributes are not supported by the readers and therefore cannot be tested:
     * $columnDimension->getAutoSize() -> false
     * $columnDimension->getCollapsed() -> true
     * $columnDimension->getColumnIndex() -> 1
     * $columnDimension->getVisible() -> false
     * $defaultColumnDimension->getAutoSize() -> true
     * $defaultColumnDimension->getCollapsed() -> false
     * $defaultColumnDimension->getColumnIndex() -> 1
     * $defaultColumnDimension->getVisible() -> true
     * $defaultRowDimension->getCollapsed() -> false
     * $defaultRowDimension->getRowIndex() -> 1
     * $defaultRowDimension->getVisible() -> true
     * $defaultRowDimension->getzeroHeight() -> false
     * $rowDimension->getCollapsed() -> true
     * $rowDimension->getRowIndex() -> 1
     * $rowDimension->getVisible() -> false
     * $rowDimension->getzeroHeight() -> true
     * $sheet->getShowGridlines() -> false.
     *
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testSheetProperties(string $format): void
    {
        $document = $this->getDocument('sheetProperties', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        $defaultColumnDimension = $sheet->getDefaultColumnDimension();
        static::assertNotNull($defaultColumnDimension, 'DefaultColumnDimension does not exist');
        static::assertSame(0, $defaultColumnDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
        static::assertSame(-1.0, $defaultColumnDimension->getWidth(), 'Unexpected value in width');
        static::assertSame(0, $defaultColumnDimension->getXfIndex(), 'Unexpected value in xfIndex');

        $columnDimension = $sheet->getColumnDimension('D');
        static::assertNotNull($columnDimension, 'ColumnDimension does not exist');
        static::assertSame(0, $columnDimension->getXfIndex(), 'Unexpected value in xfIndex');

        $pageSetup = $sheet->getPageSetup();
        static::assertNotNull($pageSetup, 'PageSetup does not exist');
        static::assertSame(1, $pageSetup->getFitToHeight(), 'Unexpected value in fitToHeight');
        static::assertFalse($pageSetup->getFitToPage(), 'Unexpected value in fitToPage');
        static::assertSame(1, $pageSetup->getFitToWidth(), 'Unexpected value in fitToWidth');
        static::assertFalse($pageSetup->getHorizontalCentered(), 'Unexpected value in horizontalCentered');
        static::assertSame(100, $pageSetup->getScale(), 'Unexpected value in scale');
        static::assertFalse($pageSetup->getVerticalCentered(), 'Unexpected value in verticalCentered');

        $defaultRowDimension = $sheet->getDefaultRowDimension();
        static::assertNotNull($defaultRowDimension, 'DefaultRowDimension does not exist');
        static::assertSame(0, $defaultRowDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
        static::assertSame(-1, $defaultRowDimension->getRowHeight(), 'Unexpected value in rowHeight');
        static::assertSame(0, (int) $defaultRowDimension->getXfIndex(), 'Unexpected value in xfIndex');
    }
}
