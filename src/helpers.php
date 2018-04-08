<?php

use PhpOffice\PhpWord\IOFactory;

if ( !function_exists('word_to_text') ) {
    /**
     * 从 word 文件中获取纯本文文字
     *
     * @param $file
     *
     * @return string
     */
    function word_to_text($file)
    {
        return strip_tags(word_to_html($file));
    }
}

if ( !function_exists('word_to_html') ) {
    /**
     * 从 word 文件中获取带样式的文本
     *
     * @param $file
     *
     * @return mixed
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    function word_to_html($file)
    {
        if (!File::isFile($file) || !in_array(File::extension($file), ['doc', 'docx'])) {
            return '';
        }
        $phpWord = IOFactory::load($file);
        $xmlWriter = IOFactory::createWriter($phpWord , 'HTML');
        $html = $xmlWriter->getWriterPart('Body')->write();
        $html = str_replace(['<body>', '</body>'], '', $html);
        return $html;
    }
}