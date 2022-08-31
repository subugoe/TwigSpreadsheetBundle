<?php

namespace MewesK\TwigSpreadsheetBundle\Tests\Twig;

class XlsxTwigTest extends BaseTwigTest
{
    public function formatProvider(): array
    {
        return [['xlsx']];
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testCellProperties(string $format): void
    {
        $document = $this->getDocument('cellProperties', $format);
        $sheet = $document->getSheetByName('Test');
        $cell = $sheet->getCell('A1');
        $dataValidation = $cell->getDataValidation();

        static::assertTrue($dataValidation->getAllowBlank(), 'Unexpected value in allowBlank');
        static::assertSame('Test error', $dataValidation->getError(), 'Unexpected value in error');
        static::assertSame('Test errorTitle', $dataValidation->getErrorTitle(), 'Unexpected value in errorTitle');
        static::assertSame('', $dataValidation->getFormula1(), 'Unexpected value in formula1');
        static::assertSame('', $dataValidation->getFormula2(), 'Unexpected value in formula2');
        static::assertSame('', $dataValidation->getOperator(), 'Unexpected value in operator');
        static::assertSame('Test prompt', $dataValidation->getPrompt(), 'Unexpected value in prompt');
        static::assertSame('Test promptTitle', $dataValidation->getPromptTitle(), 'Unexpected value in promptTitle');
        static::assertTrue($dataValidation->getShowDropDown(), 'Unexpected value in showDropDown');
        static::assertTrue($dataValidation->getShowErrorMessage(), 'Unexpected value in showErrorMessage');
        static::assertTrue($dataValidation->getShowInputMessage(), 'Unexpected value in showInputMessage');
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
        $properties = $document->getProperties();

        static::assertSame('Test company', $properties->getCompany(), 'Unexpected value in company');
        static::assertSame('Test manager', $properties->getManager(), 'Unexpected value in manager');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentTemplate(string $format): void
    {
        $document = $this->getDocument('documentTemplateAdvanced', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheet(0);
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello2', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('Bar2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');

        static::assertTrue($sheet->getCell('A1')->getStyle()->getFont()->getBold(), 'Unexpected value in bold');
        static::assertTrue($sheet->getCell('B1')->getStyle()->getFont()->getItalic(), 'Unexpected value in italic');
        static::assertSame('single', $sheet->getCell('A2')->getStyle()->getFont()->getUnderline(), 'Unexpected value in underline');
        static::assertSame('FFFF3333', $sheet->getCell('B2')->getStyle()->getFont()->getColor()->getARGB(), 'Unexpected value in color');

        $headerFooter = $sheet->getHeaderFooter();
        static::assertNotNull($headerFooter, 'HeaderFooter does not exist');
        static::assertStringContainsString('Left area header', $headerFooter->getOddHeader(),
            'Unexpected value in oddHeader');
        static::assertStringContainsString('12Center area header', $headerFooter->getOddHeader(),
            'Unexpected value in oddHeader');
        static::assertStringContainsString('12Right area header', $headerFooter->getOddHeader(),
            'Unexpected value in oddHeader');
        static::assertStringContainsString('Left area footer', $headerFooter->getOddFooter(),
            'Unexpected value in oddFooter');
        static::assertStringContainsString('12Center area footer', $headerFooter->getOddFooter(),
            'Unexpected value in oddFooter');
        static::assertStringContainsString('12Right area footer', $headerFooter->getOddFooter(),
            'Unexpected value in oddFooter');

        $drawings = $sheet->getDrawingCollection();
        static::assertCount(1, $drawings, 'Not enough drawings exist');

        $drawing = $drawings[0];
        static::assertSame(196, $drawing->getWidth(), 'Unexpected value in width');
        static::assertSame(187, $drawing->getHeight(), 'Unexpected value in height');
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
        static::assertCount(1, $drawings, 'Not enough drawings exist');

        $drawing = $drawings[0];
        static::assertSame('Test Description', $drawing->getDescription(), 'Unexpected value in description');
        static::assertSame('Test Name', $drawing->getName(), 'Unexpected value in name');
        static::assertSame(30, $drawing->getOffsetX(), 'Unexpected value in offsetX');
        static::assertSame(20, $drawing->getOffsetY(), 'Unexpected value in offsetY');
        static::assertSame(45, $drawing->getRotation(), 'Unexpected value in rotation');

        $shadow = $drawing->getShadow();
        static::assertSame('ctr', $shadow->getAlignment(), 'Unexpected value in alignment');
        static::assertSame(100, $shadow->getAlpha(), 'Unexpected value in alpha');
        static::assertSame(11, $shadow->getBlurRadius(), 'Unexpected value in blurRadius');
        static::assertSame('0000cc', $shadow->getColor()->getRGB(), 'Unexpected value in color');
        static::assertSame(30, $shadow->getDirection(), 'Unexpected value in direction');
        static::assertSame(4, $shadow->getDistance(), 'Unexpected value in distance');
        static::assertTrue($shadow->getVisible(), 'Unexpected value in visible');
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
        static::assertSame('&LfirstHeader left&CfirstHeader center&RfirstHeader right',
            $headerFooter->getFirstHeader(),
            'Unexpected value in firstHeader');
        static::assertSame('&LevenHeader left&CevenHeader center&RevenHeader right',
            $headerFooter->getEvenHeader(),
            'Unexpected value in evenHeader');
        static::assertSame('&LfirstFooter left&CfirstFooter center&RfirstFooter right',
            $headerFooter->getFirstFooter(),
            'Unexpected value in firstFooter');
        static::assertSame('&LevenFooter left&CevenFooter center&RevenFooter right',
            $headerFooter->getEvenFooter(),
            'Unexpected value in evenFooter');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testHeaderFooterDrawing(string $format): void
    {
        $document = $this->getDocument('headerFooterDrawing', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        $headerFooter = $sheet->getHeaderFooter();
        static::assertNotNull($headerFooter, 'HeaderFooter does not exist');
        static::assertSame('&L&G&CHeader', $headerFooter->getFirstHeader(), 'Unexpected value in firstHeader');
        static::assertSame('&L&G&CHeader', $headerFooter->getEvenHeader(), 'Unexpected value in evenHeader');
        static::assertSame('&L&G&CHeader', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
        static::assertSame('&LFooter&R&G', $headerFooter->getFirstFooter(), 'Unexpected value in firstFooter');
        static::assertSame('&LFooter&R&G', $headerFooter->getEvenFooter(), 'Unexpected value in evenFooter');
        static::assertSame('&LFooter&R&G', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');

        $drawings = $headerFooter->getImages();
        static::assertCount(2, $drawings, 'Sheet has not exactly 2 drawings');
        static::assertArrayHasKey('LH', $drawings, 'Header drawing does not exist');
        static::assertArrayHasKey('RF', $drawings, 'Footer drawing does not exist');

        $drawing = $drawings['LH'];
        static::assertNotNull($drawing, 'Header drawing is null');
        static::assertSame(40, $drawing->getWidth(), 'Unexpected value in width');
        static::assertSame(40, $drawing->getHeight(), 'Unexpected value in height');

        $drawing = $drawings['RF'];
        static::assertNotNull($drawing, 'Footer drawing is null');
        static::assertSame(20, $drawing->getWidth(), 'Unexpected value in width');
        static::assertSame(20, $drawing->getHeight(), 'Unexpected value in height');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testHeaderFooterProperties(string $format): void
    {
        $document = $this->getDocument('headerFooterProperties', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheetByName('Test');
        static::assertNotNull($sheet, 'Sheet does not exist');

        $headerFooter = $sheet->getHeaderFooter();
        static::assertNotNull($headerFooter, 'HeaderFooter does not exist');

        static::assertSame('&CHeader', $headerFooter->getFirstHeader(), 'Unexpected value in firstHeader');
        static::assertSame('&CHeader', $headerFooter->getEvenHeader(), 'Unexpected value in evenHeader');
        static::assertSame('&CHeader', $headerFooter->getOddHeader(), 'Unexpected value in oddHeader');
        static::assertSame('&CFooter', $headerFooter->getFirstFooter(), 'Unexpected value in firstFooter');
        static::assertSame('&CFooter', $headerFooter->getEvenFooter(), 'Unexpected value in evenFooter');
        static::assertSame('&CFooter', $headerFooter->getOddFooter(), 'Unexpected value in oddFooter');

        static::assertFalse($headerFooter->getAlignWithMargins(), 'Unexpected value in alignWithMargins');
        static::assertFalse($headerFooter->getScaleWithDocument(), 'Unexpected value in scaleWithDocument');
    }

    /**
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
        static::assertSame('A1:B1', $sheet->getAutoFilter()->getRange(), 'Unexpected value in autoFilter');
    }
}
