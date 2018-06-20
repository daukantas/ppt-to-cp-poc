<?php

require_once './PHPPresentation/src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();

require_once './Common/src/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\DocumentLayout;

if (sizeof($argv) !== 2) {
  echo "I need the path to the pptx file\n";
  die;
}
$reader = IOFactory::createReader('PowerPoint2007');
$presentation = $reader->load($argv[1]);

function dumpShape($shape) {
  $props = [
    'type' => get_class($shape),
    'height' => $shape->getHeight(),
    'width' => $shape->getWidth(),
    'position' => [
      'x' => $shape->getOffsetX(),
      'y' => $shape->getOffsetY()
    ]
  ];

  if($shape instanceof Drawing\Gd) {
    $props['mime'] = $shape->getMimeType();
    $props['name'] = $shape->getName();
    $props['description'] = $shape->getDescription();

    ob_start();
    call_user_func($shape->getRenderingFunction(), $shape->getImageResource());
    $img = ob_get_contents();
    ob_end_clean();
    $props['content'] = base64_encode($img);
  }
  elseif($shape instanceof Drawing\File) {

  }
  elseif($shape instanceof Drawing\Base64) {

  }
  elseif($shape instanceof Drawing\ZipFile) {

  }
  elseif($shape instanceof RichText) {

    $text = '';
    foreach ($shape->getParagraphs() as $p) {
      $text .= '<p>' . $p->getPlainText() . '</p>';
    }

    $props['text'] = $text;
  }

  return $props;
}

// Get
$slideIndex = 0;
$output = [
  'slides' => [],
  'width' => $presentation->getLayout()->getCX(DocumentLayout::UNIT_PIXEL)
];
foreach ($presentation->getAllSlides() as $slide) {
  $shapes = [];

  // Slide backgrounds
  $background = ['type' => 'none'];
  $slide_bg = $slide->getBackground();
  if ($slide_bg instanceof Slide\Background\Color) {
    $background = [
      'type' => 'color',
      'rgb' => $slide_bg->getColor()->getRGB()
    ];
  }
  elseif ($slide_bg instanceof Slide\Background\Image) {
    $img = file_get_contents($slide_bg->getPath());

    $background = [
      'type' => 'image',
      'mime' => '',
      'content' => base64_encode($img)
    ];
  }

  // Shapes
  foreach ($slide->getShapeCollection() as $shape) {
    if($shape instanceof Group) {
      foreach ($shape->getShapeCollection() as $child) {
        $shapes[] = dumpShape($child);
      }
    }
    else {
      $shapes[] = dumpShape($shape);
    }
  }
  $output['slides'][] = [
    'background' => $background,
    'shapes' => $shapes
  ];

  if ($slideIndex++ > 5) {
    break;
  }
}

echo json_encode($output) . "\n\n";
