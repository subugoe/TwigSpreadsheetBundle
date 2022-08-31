<?php

namespace MewesK\TwigSpreadsheetBundle\Tests\Twig;

class CsvOdsXlsXlsxTwigTest extends BaseTwigTest
{
    public function formatProvider(): array
    {
        return [['csv'], ['ods'], ['xls'], ['xlsx']];
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testDocumentSimple(string $format): void
    {
        $document = $this->getDocument('documentSimple', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getActiveSheet();
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
    public function testDocumentTemplate(string $format): void
    {
        $document = $this->getDocument('documentTemplate.'.$format, $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheet(0);
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame('Hello2', $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame('World', $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertSame('Foo', $sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertSame('Bar2', $sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }

    /**
     * @throws \Exception
     *
     * @dataProvider formatProvider
     */
    public function testFunctionIndex(string $format): void
    {
        $document = $this->getDocument('functionIndex', $format);
        static::assertNotNull($document, 'Document does not exist');

        $sheet = $document->getSheet(0);
        static::assertNotNull($sheet, 'Sheet does not exist');

        static::assertSame(1, $sheet->getCell('A1')->getValue(), 'Unexpected value in A1');
        static::assertSame(2, $sheet->getCell('B1')->getValue(), 'Unexpected value in B1');
        static::assertNull($sheet->getCell('A2')->getValue(), 'Unexpected value in A2');
        static::assertNull($sheet->getCell('B2')->getValue(), 'Unexpected value in B2');
    }
}
