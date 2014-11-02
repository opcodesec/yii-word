<?php

/**
 * Wrapper for the PHPWord library.
 * @see README.md
 */
class YiiWord extends CComponent
{
     public static $_pathAlias = 'application.vendors.PHPWord';

    /**
     * Register autoloader.
     */
    static function autoload($pClassName){
        if ((class_exists($pClassName, false)) || (strpos($pClassName, 'PHPWord') !== 0)) {
            //  Either already loaded, or not a PHPExcel class request
            return false;
        }

        //get the path
        //$pClassFilePath = Yii::getPathOfAlias('application.vendors.phpexcel').'/'
        $pClassFilePath = Yii::getPathOfAlias(self::$_pathAlias).'/'
            .str_replace('_', DIRECTORY_SEPARATOR, $pClassName).'.php';

        if ((file_exists($pClassFilePath) === false) || (is_readable($pClassFilePath) === false)) {
            //  Can't load
            return false;
        }

        require($pClassFilePath);
    }//loadClass end
}