<?php

namespace MewesK\TwigSpreadsheetBundle\Tests\Twig;

use PhpOffice\PhpSpreadsheet\Shared\PasswordHasher;

class XlsXlsxTwigTest extends BaseTwigTest
{
    public function formatProvider(): array
    {
        return [['xls'], ['xlsx']];
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testCellIndexMerge(string $format): void
    {
        $document = $this->getDocument('cellIndexMerge', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('A2:C2', $sheet->getCell('A2')->getMergeRange(), 'Unexpected value in mergeRange');
        static::assertSame('A3:C3', $sheet->getCell('A3')->getMergeRange(), 'Unexpected value in mergeRange');
        static::assertSame('A4:A6', $sheet->getCell('A4')->getMergeRange(), 'Unexpected value in mergeRange');
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

        $breaks = $sheet->getBreaks();
        static::assertCount(1, $breaks, 'Unexpected break count');
        static::assertArrayHasKey('A1', $breaks, 'Break does not exist');

        $break = $breaks['A1'];
        static::assertNotNull($break, 'Break is null');

        $font = $cell->getStyle()->getFont();
        static::assertNotNull($font, 'Font does not exist');
        static::assertSame(18, (int) $font->getSize(), 'Unexpected value in size');

        static::assertSame('https://example.com/', $cell->getHyperlink()->getUrl(), 'Unexpected value in url');
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

        static::assertSame('Test category', $properties->getCategory(), 'Unexpected value in category');

        $font = $document->getDefaultStyle()->getFont();
        static::assertNotNull($font, 'Font does not exist');
        static::assertSame(18, (int) $font->getSize(), 'Unexpected value in size');

        static::assertSame('Test keywords', $properties->getKeywords(), 'Unexpected value in keywords');
        static::assertSame('Test modifier', $properties->getLastModifiedBy(), 'Unexpected value in lastModifiedBy');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDrawingProperties(string $format): void
    {
        $document = $this->getDocument('drawingProperties', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        $drawings = $sheet->getDrawingCollection();
        static::assertCount(1, $drawings, 'Unexpected drawing count');
        static::assertArrayHasKey(0, $drawings, 'Drawing does not exist');

        $drawing = $drawings[0];
        static::assertNotNull($drawing, 'Drawing is null');

        static::assertSame('B2', $drawing->getCoordinates(), 'Unexpected value in coordinates');
        static::assertSame(200, $drawing->getHeight(), 'Unexpected value in height');
        static::assertFalse($drawing->getResizeProportional(), 'Unexpected value in resizeProportional');
        static::assertSame(300, $drawing->getWidth(), 'Unexpected value in width');

        $shadow = $drawing->getShadow();
        static::assertNotNull($shadow, 'Shadow is null');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDrawingSimple(string $format): void
    {
        $document = $this->getDocument('drawingSimple', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        $drawings = $sheet->getDrawingCollection();
        static::assertCount(1, $drawings, 'Unexpected drawing count');
        static::assertArrayHasKey(0, $drawings, 'Drawing does not exist');

        $drawing = $drawings[0];
        static::assertNotNull($drawing, 'Drawing is null');
        static::assertSame(100, $drawing->getWidth(), 'Unexpected value in width');
        static::assertSame(100, $drawing->getHeight(), 'Unexpected value in height');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testFunctionMergeStyles(string $format): void
    {
        $document = $this->getDocument('functionMergeStyles', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheet(0);
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Calibri', $sheet->getCell('A1')->getStyle()->getFont()->getName(), 'Unexpected value in A1');
        static::assertSame(11, (int) $sheet->getCell('A1')->getStyle()->getFont()->getSize(), 'Unexpected value in A1');
        static::assertSame(11, (int) $sheet->getCell('A2')->getStyle()->getFont()->getSize(), 'Unexpected value in A2');
        static::assertSame(11, (int) $sheet->getCell('A3')->getStyle()->getFont()->getSize(), 'Unexpected value in A3');
        static::assertSame(11, (int) $sheet->getCell('A4')->getStyle()->getFont()->getSize(), 'Unexpected value in B3');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testHeaderFooterComplex(string $format): void
    {
        $document = $this->getDocument('headerFooterComplex', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        $headerFooter = $sheet->getHeaderFooter();
        static::assertNotNull($headerFooter, 'HeaderFooter does not exist');

        static::assertSame('&LoddHeader left&CoddHeader center&RoddHeader right', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
        static::assertSame('&LoddFooter left&CoddFooter center&RoddFooter right', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');
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

        $columnDimension = $sheet->getColumnDimension('D');
        static::assertSame(1, $columnDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
        static::assertSame(200, (int) $columnDimension->getWidth(), 'Unexpected value in width');

        $pageMargins = $sheet->getPageMargins();
        static::assertNotNull($pageMargins, 'PageMargins does not exist');
        static::assertSame(1, (int) $pageMargins->getTop(), 'Unexpected value in top');
        static::assertSame(1, (int) $pageMargins->getBottom(), 'Unexpected value in bottom');
        static::assertSame(0.75, $pageMargins->getLeft(), 'Unexpected value in left');
        static::assertSame(0.75, $pageMargins->getRight(), 'Unexpected value in right');
        static::assertSame(0.5, $pageMargins->getHeader(), 'Unexpected value in header');
        static::assertSame(0.5, $pageMargins->getFooter(), 'Unexpected value in footer');

        $pageSetup = $sheet->getPageSetup();
        static::assertSame('landscape', $pageSetup->getOrientation(), 'Unexpected value in orientation');
        static::assertSame(9, $pageSetup->getPaperSize(), 'Unexpected value in paperSize');
        static::assertSame('A1:B1', $pageSetup->getPrintArea(), 'Unexpected value in printArea');

        $protection = $sheet->getProtection();
        static::assertTrue($protection->getAutoFilter(), 'Unexpected value in autoFilter');
        static::assertNotNull($protection, 'Protection does not exist');
        static::assertTrue($protection->getDeleteColumns(), 'Unexpected value in deleteColumns');
        static::assertTrue($protection->getDeleteRows(), 'Unexpected value in deleteRows');
        static::assertTrue($protection->getFormatCells(), 'Unexpected value in formatCells');
        static::assertTrue($protection->getFormatColumns(), 'Unexpected value in formatColumns');
        static::assertTrue($protection->getFormatRows(), 'Unexpected value in formatRows');
        static::assertTrue($protection->getInsertColumns(), 'Unexpected value in insertColumns');
        static::assertTrue($protection->getInsertHyperlinks(), 'Unexpected value in insertHyperlinks');
        static::assertTrue($protection->getInsertRows(), 'Unexpected value in insertRows');
        static::assertTrue($protection->getObjects(), 'Unexpected value in objects');
        static::assertSame(PasswordHasher::hashPassword('testpassword'), $protection->getPassword(), 'Unexpected value in password');
        static::assertTrue($protection->getPivotTables(), 'Unexpected value in pivotTables');
        static::assertTrue($protection->getScenarios(), 'Unexpected value in scenarios');
        static::assertTrue($protection->getSelectLockedCells(), 'Unexpected value in selectLockedCells');
        static::assertTrue($protection->getSelectUnlockedCells(), 'Unexpected value in selectUnlockedCells');
        static::assertTrue($protection->getSheet(), 'Unexpected value in sheet');
        static::assertTrue($protection->getSort(), 'Unexpected value in sort');

        static::assertTrue($sheet->getPrintGridlines(), 'Unexpected value in printGridlines');
        static::assertTrue($sheet->getRightToLeft(), 'Unexpected value in rightToLeft');
        static::assertSame('c0c0c0', strtolower($sheet->getTabColor()->getRGB()), 'Unexpected value in tabColor');
        static::assertSame(75, $sheet->getSheetView()->getZoomScale(), 'Unexpected value in zoomScale');

        $rowDimension = $sheet->getRowDimension(2);
        static::assertNotNull($rowDimension, 'RowDimension does not exist');
        static::assertSame(1, $rowDimension->getOutlineLevel(), 'Unexpected value in outlineLevel');
        static::assertSame(30, (int) $rowDimension->getRowHeight(), 'Unexpected value in rowHeight');
        static::assertSame(0, (int) $rowDimension->getXfIndex(), 'Unexpected value in xfIndex');
    }
}
