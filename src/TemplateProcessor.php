<?php

namespace Mnvx\EloquentPrintForm;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Arr;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpWord\Exception\Exception;

class TemplateProcessor extends \PhpOffice\PhpWord\TemplateProcessor
{
    private $currentPart = null;
    public function cloneRow($search, $numberOfClones)
    {
        $this->currentPart = $this->tempDocumentMainPart;
        $this->cloneRowInPart($this->tempDocumentMainPart, $search, $numberOfClones);

        foreach ($this->tempDocumentFooters as &$footer) {
            $this->currentPart = $footer;
            $this->cloneRowInPart($footer, $search, $numberOfClones);
        }

        foreach ($this->tempDocumentHeaders as &$header) {
            $this->currentPart = $header;
            $this->cloneRowInPart($header, $search, $numberOfClones);
        }
    }

    public function cloneRowInPart(&$part, $search, $numberOfClones)
    {
        $search = static::ensureMacroCompleted($search);

        $tagPos = strpos($part, $search);
        if (!$tagPos) {
            return;
//            dd($part, $search);
            throw new Exception('Can not clone row, template variable not found or variable contains markup.');
        }

        $rowStart = $this->findRowStart($tagPos, $part);
        $rowEnd = $this->findRowEnd($tagPos, $part);
        $xmlRow = $this->getSlice($rowStart, $rowEnd, $part);

        // Check if there's a cell spanning multiple rows.
        if (preg_match('#<w:vMerge w:val="restart"/>#', $xmlRow)) {
            // $extraRowStart = $rowEnd;
            $extraRowEnd = $rowEnd;
            while (true) {
                $extraRowStart = $this->findRowStart($extraRowEnd + 1);
                $extraRowEnd = $this->findRowEnd($extraRowEnd + 1);

                // If extraRowEnd is lower then 7, there was no next row found.
                if ($extraRowEnd < 7) {
                    break;
                }

                // If tmpXmlRow doesn't contain continue, this row is no longer part of the spanned row.
                $tmpXmlRow = $this->getSlice($extraRowStart, $extraRowEnd);
                if (!preg_match('#<w:vMerge/>#', $tmpXmlRow) &&
                    !preg_match('#<w:vMerge w:val="continue"\s*/>#', $tmpXmlRow)) {
                    break;
                }
                // This row was a spanned row, update $rowEnd and search for the next row.
                $rowEnd = $extraRowEnd;
            }
            $xmlRow = $this->getSlice($rowStart, $rowEnd);
        }

        $result = $this->getSlice(0, $rowStart);
        $result .= implode($this->indexClonedVariables($numberOfClones, $xmlRow));
        $result .= $this->getSlice($rowEnd);

        $part = $result;
    }
//
//    public function setComplexValueInPart($search, \PhpOffice\PhpWord\Element\AbstractElement $complexType)
//    {
//        $elementName = substr(get_class($complexType), strrpos(get_class($complexType), '\\') + 1);
//        $objectClass = 'PhpOffice\\PhpWord\\Writer\\Word2007\\Element\\' . $elementName;
//
//        $xmlWriter = new XMLWriter();
//        /** @var \PhpOffice\PhpWord\Writer\Word2007\Element\AbstractElement $elementWriter */
//        $elementWriter = new $objectClass($xmlWriter, $complexType, true);
//        $elementWriter->write();
//
//        $where = $this->findContainingXmlBlockForMacro($search, 'w:r');
//        $block = $this->getSlice($where['start'], $where['end']);
//        $textParts = $this->splitTextIntoTexts($block);
//        $this->replaceXmlBlock($search, $textParts, 'w:r');
//
//        $search = static::ensureMacroCompleted($search);
//        $this->replaceXmlBlock($search, $xmlWriter->getData(), 'w:r');
//    }

    public function setComplexValue($search, \PhpOffice\PhpWord\Element\AbstractElement $complexType)
    {
        $this->currentPart = $this->tempDocumentMainPart;
        $this->setComplexValueLocal($search, $complexType);

        foreach ($this->tempDocumentFooters as &$footer) {
            $this->currentPart = $footer;
            $this->setComplexValueLocal($search, $complexType);
        }

        foreach ($this->tempDocumentHeaders as &$header) {
            $this->currentPart = $header;
            $this->setComplexValueLocal($search, $complexType);
        }
//        dd();
    }

    /**
     * @param string $search
     * @param \PhpOffice\PhpWord\Element\AbstractElement $complexType
     */
    public function setComplexValueLocal($search, \PhpOffice\PhpWord\Element\AbstractElement $complexType)
    {
        $elementName = substr(get_class($complexType), strrpos(get_class($complexType), '\\') + 1);
        $objectClass = 'PhpOffice\\PhpWord\\Writer\\Word2007\\Element\\' . $elementName;

        $xmlWriter = new XMLWriter();
        /** @var \PhpOffice\PhpWord\Writer\Word2007\Element\AbstractElement $elementWriter */
        $elementWriter = new $objectClass($xmlWriter, $complexType, true);
        $elementWriter->write();

        $where = $this->findContainingXmlBlockForMacro($search, 'w:r');

        if (!$where) {
            return;
        }

        $block = $this->getSlice($where['start'], $where['end']);
        $textParts = $this->splitTextIntoTexts($block);
//        dump($this->currentPart, $search, $textParts);
        $this->replaceXmlBlock($search, $textParts, 'w:r');
        $search = static::ensureMacroCompleted($search);
        $this->replaceXmlBlock($search, $xmlWriter->getData(), 'w:r');
    }

        /**
         * Find the start position of the nearest table row before $offset.
         *
         * @param int $offset
         *
         * @throws \PhpOffice\PhpWord\Exception\Exception
         *
         * @return int
         */
        protected function findRowStart($offset)
        {
            $part = $this->currentPart ?? $this->tempDocumentMainPart;
            $rowStart = strrpos($part, '<w:tr ', ((strlen($part) - $offset) * -1));

            if (!$rowStart) {
                $rowStart = strrpos($part, '<w:tr>', ((strlen($part) - $offset) * -1));
            }
            if (!$rowStart) {
                throw new Exception('Can not find the start position of the row to clone.');
            }

            return $rowStart;
        }

        /**
         * Find the end position of the nearest table row after $offset.
         *
         * @param int $offset
         *
         * @return int
         */
        protected function findRowEnd($offset)
        {
            $part = $this->currentPart ?? $this->tempDocumentMainPart;
            return strpos($part, '</w:tr>', $offset) + 7;
        }

        /**
         * Get a slice of a string.
         *
         * @param int $startPosition
         * @param int $endPosition
         *
         * @return string
         */
        protected function getSlice($startPosition, $endPosition = 0)
        {
            $part = $this->currentPart ?? $this->tempDocumentMainPart;
            if (!$endPosition) {
                $endPosition = strlen($part);
            }

            return substr($part, $startPosition, ($endPosition - $startPosition));
        }

    /**
     * Find the position of (the start of) a macro
     *
     * Returns -1 if not found, otherwise position of opening $
     *
     * Note that only the first instance of the macro will be found
     *
     * @param string $search Macro name
     * @param int $offset Offset from which to start searching
     * @return int -1 if macro not found
     */
    protected function findMacro($search, $offset = 0)
    {
        $part = $this->currentPart ?? $this->tempDocumentMainPart;
        $search = static::ensureMacroCompleted($search);
        $pos = strpos($part, $search, $offset);

        return ($pos === false) ? -1 : $pos;
    }

    /**
     * Find the start position of the nearest XML block start before $offset
     *
     * @param int $offset    Search position
     * @param string  $blockType XML Block tag
     * @return int -1 if block start not found
     */
    protected function findXmlBlockStart($offset, $blockType)
    {
        $part = $this->currentPart ?? $this->tempDocumentMainPart;
        $reverseOffset = (strlen($part) - $offset) * -1;
        // first try XML tag with attributes
        $blockStart = strrpos($part, '<' . $blockType . ' ', $reverseOffset);
        // if not found, or if found but contains the XML tag without attribute
        if (false === $blockStart || strrpos($this->getSlice($blockStart, $offset), '<' . $blockType . '>')) {
            // also try XML tag without attributes
            $blockStart = strrpos($part, '<' . $blockType . '>', $reverseOffset);
        }

        return ($blockStart === false) ? -1 : $blockStart;
    }

    /**
     * Find the nearest block end position after $offset
     *
     * @param int $offset    Search position
     * @param string  $blockType XML Block tag
     * @return int -1 if block end not found
     */
    protected function findXmlBlockEnd($offset, $blockType)
    {
        $part = $this->currentPart ?? $this->tempDocumentMainPart;
        $blockEndStart = strpos($part, '</' . $blockType . '>', $offset);
        // return position of end of tag if found, otherwise -1

        return ($blockEndStart === false) ? -1 : $blockEndStart + 3 + strlen($blockType);
    }
//    /**
//     * Replace an XML block surrounding a macro with a new block
//     *
//     * @param string $macro Name of macro
//     * @param string $block New block content
//     * @param string $blockType XML tag type of block
//     * @return \PhpOffice\PhpWord\TemplateProcessor Fluent interface
//     */
    public function replaceXmlBlock($macro, $block, $blockType = 'w:p')
    {
        $part = $this->tempDocumentMainPart;
//        dd($part);

            $this->currentPart = $this->tempDocumentMainPart;
            $where = $this->findContainingXmlBlockForMacro($macro, $blockType);

            if (is_array($where)) {
                $this->tempDocumentMainPart = $this->getSlice(0, $where['start']) . $block . $this->getSlice($where['end']);
            }

            foreach ($this->tempDocumentFooters as &$footer) {
                $this->currentPart = $footer;
                $where = $this->findContainingXmlBlockForMacro($macro, $blockType);
                if (is_array($where)) {
                    $footer = $this->getSlice(0, $where['start']) . $block . $this->getSlice($where['end']);
                }
            }

            foreach ($this->tempDocumentHeaders as &$header) {
                $this->currentPart = $header;
                $where = $this->findContainingXmlBlockForMacro($macro, $blockType);
                if (is_array($where)) {
                    $header = $this->getSlice(0, $where['start']) . $block . $this->getSlice($where['end']);
                }
            }
        return $this;
    }
}
