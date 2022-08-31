<?php

namespace MewesK\TwigSpreadsheetBundle\Wrapper;

use InvalidArgumentException;
use LogicException;
use MewesK\TwigSpreadsheetBundle\Helper\Filesystem;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\BaseWriter;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;
use RuntimeException;
use Symfony\Bridge\Twig\AppVariable;
use Symfony\Component\Filesystem\Exception\IOException;
use Twig\Environment;
use Twig\Loader\FilesystemLoader;

class DocumentWrapper extends BaseWrapper
{
    protected ?Spreadsheet $object = null;

    public function __construct(array $context, Environment $environment, protected array $attributes = [])
    {
        parent::__construct($context, $environment);
    }

    /**
     * @throws RuntimeException
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function start(array $properties = []): void
    {
        // load template
        if (isset($properties['template'])) {
            $templatePath = $this->expandPath($properties['template']);
            $reader = IOFactory::createReaderForFile($templatePath);
            $this->object = $reader->load($templatePath);
        }

        // create new
        else {
            $this->object = new Spreadsheet();
            $this->object->removeSheetByIndex(0);
        }

        $this->parameters['properties'] = $properties;

        $this->setProperties($properties);
    }

    /**
     * @throws LogicException
     * @throws RuntimeException
     * @throws InvalidArgumentException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     * @throws IOException
     */
    public function end(): void
    {
        if (null === $this->object) {
            throw new LogicException();
        }

        $format = null;

        // try document property
        if (isset($this->parameters['format'])) {
            $format = $this->parameters['format'];
        }

        // try Symfony request
        elseif (isset($this->context['app'])) {
            /**
             * @var AppVariable $appVariable
             */
            $appVariable = $this->context['app'];
            if ($appVariable instanceof AppVariable && null !== $appVariable->getRequest()) {
                $format = $appVariable->getRequest()->getRequestFormat();
            }
        }

        // set default
        $format = \is_string($format) ? strtolower($format) : 'xlsx';

        // set up mPDF
        if ('pdf' === $format) {
            if (!class_exists(\Mpdf\Mpdf::class)) {
                throw new RuntimeException('Error loading mPDF. Is mPDF correctly installed?');
            }

            IOFactory::registerWriter('Pdf', Mpdf::class);
        }

        /**
         * @var BaseWriter $writer
         */
        $writer = IOFactory::createWriter($this->object, ucfirst($format));
        $writer->setPreCalculateFormulas($this->attributes['pre_calculate_formulas'] ?? true);

        // set up XML cache
        if (false !== $this->attributes['cache']['xml']) {
            Filesystem::mkdir($this->attributes['cache']['xml']);
            $writer->setUseDiskCaching(true, $this->attributes['cache']['xml']);
        }

        // set special CSV writer attributes
        if ($writer instanceof Csv) {
            /*
             * @var Csv $writer
             */
            $writer->setDelimiter($this->attributes['csv_writer']['delimiter']);
            $writer->setEnclosure($this->attributes['csv_writer']['enclosure']);
            $writer->setExcelCompatibility($this->attributes['csv_writer']['excel_compatibility']);
            $writer->setIncludeSeparatorLine($this->attributes['csv_writer']['include_separator_line']);
            $writer->setLineEnding($this->attributes['csv_writer']['line_ending']);
            $writer->setSheetIndex($this->attributes['csv_writer']['sheet_index']);
            $writer->setUseBOM($this->attributes['csv_writer']['use_bom']);
        }

        $writer->save('php://output');

        $this->object = null;
        $this->parameters = [];
    }

    public function getObject(): ?Spreadsheet
    {
        return $this->object;
    }

    public function setObject(Spreadsheet $object = null): void
    {
        $this->object = $object;
    }

    /**
     * {@inheritdoc}
     */
    protected function configureMappings(): array
    {
        return [
            'category' => function ($value) { $this->object->getProperties()->setCategory($value); },
            'company' => function ($value) { $this->object->getProperties()->setCompany($value); },
            'created' => function ($value) { $this->object->getProperties()->setCreated($value); },
            'creator' => function ($value) { $this->object->getProperties()->setCreator($value); },
            'defaultStyle' => function ($value) { $this->object->getDefaultStyle()->applyFromArray($value); },
            'description' => function ($value) { $this->object->getProperties()->setDescription($value); },
            'format' => function ($value) { $this->parameters['format'] = $value; },
            'keywords' => function ($value) { $this->object->getProperties()->setKeywords($value); },
            'lastModifiedBy' => function ($value) { $this->object->getProperties()->setLastModifiedBy($value); },
            'manager' => function ($value) { $this->object->getProperties()->setManager($value); },
            'modified' => function ($value) { $this->object->getProperties()->setModified($value); },
            'security' => [
                'lockRevision' => function ($value) { $this->object->getSecurity()->setLockRevision($value); },
                'lockStructure' => function ($value) { $this->object->getSecurity()->setLockStructure($value); },
                'lockWindows' => function ($value) { $this->object->getSecurity()->setLockWindows($value); },
                'revisionsPassword' => function ($value) { $this->object->getSecurity()->setRevisionsPassword($value); },
                'workbookPassword' => function ($value) { $this->object->getSecurity()->setWorkbookPassword($value); },
            ],
            'subject' => function ($value) { $this->object->getProperties()->setSubject($value); },
            'template' => function ($value) { $this->parameters['template'] = $value; },
            'title' => function ($value) { $this->object->getProperties()->setTitle($value); },
        ];
    }

    /**
     * Resolves paths using Twig namespaces.
     * The path must start with the namespace.
     * Namespaces are case sensitive.
     */
    private function expandPath(string $path): string
    {
        $loader = $this->environment->getLoader();

        if ($loader instanceof FilesystemLoader && str_starts_with($path, '@')) {
            /*
             * @var \Twig\Loader\FilesystemLoader
             */
            foreach ($loader->getNamespaces() as $namespace) {
                if (1 === mb_strpos($path, $namespace)) {
                    foreach ($loader->getPaths($namespace) as $namespacePath) {
                        $expandedPathAttribute = str_replace('@'.$namespace, $namespacePath, $path);
                        if (Filesystem::exists($expandedPathAttribute)) {
                            return $expandedPathAttribute;
                        }
                    }
                }
            }
        }

        return $path;
    }
}
