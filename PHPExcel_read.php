<?
//2007 xlsx格式
public function readOnlyExcel($file,$type='Excel2007')
{
    $result = array();
    $objReader = \PHPExcel_IOFactory::createReader($type);
    $objReader->setReadDataOnly(TRUE);
    $objPHPExcel  = $objReader->load($file);           //载入Excel文件 

    $sheet	            = $objPHPExcel->getSheet(0);  //读取excel文件中的第一个工作表
    $highestRow			= $sheet->getHighestRow();    //取得一共有多少行
    $highestColumn		= $sheet->getHighestColumn();     //取得最大的列号
    $highestColumnIndex	= \PHPExcel_Cell::columnIndexFromString($highestColumn);//字母列转换为数字列 如:AA变为27

    /** 循环读取每个单元格的数据 */
    for($i=($type=='Excel2007'?1:2);$i<=$highestRow;$i++)      //行数是以第1行开始
    {
        $row = array();
        for($k=0;$k<$highestColumnIndex;$k++)           //列数是以第0列开始
        {
            $v = $sheet->getCellByColumnAndRow($k,$i)->getValue();//读取单元格
            if(is_object($v))
            {
                array_push($row,'');
                continue;
            }
            array_push($row,$v);
        }
        array_push($result,$row);
    }

    return $result;
}