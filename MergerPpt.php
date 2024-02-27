<?php

require_once 'src/PhpPresentation/src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();

require_once 'src/Common/src/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;

class MergerPpt
{
    protected PhpPresentation $presentation;

    public function __construct(PhpPresentation $pres)
    {
        $this->presentation = $pres;
    }

    public function merge(string $ppt1, string $ppt2, string $outputFileName)
    {
        $p1 = $this->read($ppt1);
        $p2 = $this->read($ppt2);

        $this->add($p1->getAllSlides());
        $this->add($p2->getAllSlides());

        $this->slideRemove(0);

        $this->save($outputFileName);
    }

    private function read(string $fileName): PhpPresentation
    {
        $reader = IOFactory::createReader('PowerPoint2007');
        return $reader->load($fileName);
    }

    private function add(array $slides)
    {
        for($i=0; $i<count($slides); $i++){
            $this->presentation->addSlide($slides[$i]);
        }
    }

    private function slideRemove(int $index): PhpPresentation
    {
        return $this->presentation->removeSlideByIndex($index);
    }

    public function slideRemoveFromPresentation(PhpPresentation $presentation ,int $index): PhpPresentation
    {
        return $presentation->removeSlideByIndex($index);
    }

    private function save(string $fileName)
    {
        $writer = IOFactory::createWriter($this->presentation, 'PowerPoint2007');
        $writer->save($fileName);
    }
}

function clientCode(string $ppt1, string $ppt2, string $outputFileName){
    $presentation = new PhpPresentation();
    $merger = new MergerPpt($presentation);
    $merger->merge($ppt1, $ppt2, $outputFileName);
}

clientCode("pp1.pptx", "pp1.pptx", "test.pptx");
