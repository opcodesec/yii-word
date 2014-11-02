yii-word
========

Word extension for YII (depend on the [PHPWord library](https://phpword.codeplex.com/))

### Install ###

1. git clone yii-word from github;
2. mkdir a directory (named: yiiword) under protected/extensions;
3. Unzip the YiiWord to protected/extensions/yiiword;
4. Download the latest version of PHPWord: https://phpword.codeplex.com/;
5. Unzip the contents of the folder Classes to a new folder protected/vendor/PHPWord/;
6. import this library "application.vendors.PHPWord.PHPWord" to the main.php;

### Usage ###

    Yii::import('ext.yiiword.YiiWord', true);
    Yii::registerAutoloader(array('YiiWord', 'autoload'), true);

    $PHPWord = new PHPWord();
    $section = $PHPWord->createSection();

    $section->addText('Hell Wordl!');

    //generate download page
    $filename = time();
    header('Content-Type: application/vnd.ms-word');
    header('Content-Disposition: attachment;filename="'.$filename.'.docx"');
    header('Cache-Control: max-age=0');

    $objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
    $objWriter->save('php://output');
    unset($this->objWriter);
    exit();

### Refrence ###

1. [phpword](https://phpword.codeplex.com/);
2. [yii-phpword](https://github.com/websthetics/yii-phpword);
