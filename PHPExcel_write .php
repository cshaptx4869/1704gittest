//导出
<?php
 $dir = dirname(__FILE__);  //找出当前脚本所在路径
 require $dir.'\lib\PHPExcel_1.8.0_doc\Classes\PHPExcel.php'; //添加读取excel所需的类文件

 $objPHPExcel = new PHPExcel();                     //实例化一个PHPExcel()对象
 $objSheet = $objPHPExcel->getActiveSheet();        //选取当前的sheet对象
 $objSheet->setTitle('helen');                      //对当前sheet对象命名
 //常规方式：利用setCellValue()填充数据
 $objSheet->setCellValue("A1","张三")->setCellValue("B1","李四");   //利用setCellValues()填充数据
 //取巧模式：利用fromArray()填充数据
 $array = array(
     array("","B1","张三"),
     array("","B2","李四")
 );
 $objSheet->fromArray($array);  //利用fromArray()直接一次性填充数据
 
 $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');   //设定写入excel的类型
 $objWriter->save($dir.'/test.xlsx');      //保存文件
 ?>

//PHP从数据库导出EXCEL文件

<!-- 参考博客链接：http://www.cnblogs.com/huangcong/p/3687665.html

我的程序代码

原生导出Excel文件 -->

<?php
header('Content-type: text/html; charset=utf-8');
header("Content-type:application/vnd.ms-excel");
header("Content-Disposition:filename=test.xls");

$conn = mysqli_connect("localhost","zhouqi","445864742") or die("无法连接数据库");
mysqli_select_db($conn,"test");
mysqli_set_charset($conn,'utf8');
$sql = "SELECT * FROM student";
$result = mysqli_query($conn,$sql);
echo "ID号\t姓名\t分数\t\n";
while ($row = mysqli_fetch_array($result)){
    echo $row[0]."\t".$row[1]."\t".$row[2]."\t\n";
}
?>

 

<!-- \t为换格   \n为换行

 

 

PHPEXCEL用法 -->

 

 

 

<?php
require_once(phpexcel_dir());	//引入PHPExcel 类

$objPHPExcel=new PHPExcel();
//获得数据  ---一般是从数据库中获得数据

$conn = mysqli_connect(Conf::$db_host,Conf::$db_username,Conf::$db_password) or die("无法连接数据库");//连接数据库主机，用户名，密码 配置文件里设置
mysqli_select_db($conn,Conf::$db_dbname);//选择数据库
mysqli_set_charset($conn,'utf8');	//设置字符集

 

 

 

//左连接连接三张表
$sql = "SELECT 
              a.id,a.order_sn,a.status,a.rev_name,a.rev_addr,a.rev_mail,a.rev_post,a.rev_mobile,b.account_name,c.brands_name,a.project 
            FROM ec_orders AS a 
            LEFT JOIN ec_account AS b ON a.account_id = b.id

            LEFT JOIN ec_goods_brands AS c ON a.brands_id = c.id
            WHERE a.is_del = 0 AND b.is_del = 0 AND c.is_del =0";
$result = mysqli_query($conn,$sql);
$data = array();
$i = 0;
while ($row = mysqli_fetch_array($result)){
    $data[$i]['id'] = $row['id'];
    $data[$i]['order_sn'] = $row['order_sn'];
    //订单的状态   0:待确认 1:已确认/待付款 2:已付款/待发货 3:发货中 4:已发货 5:买家收货确认 6:订单完成 7:买家取消订单 8:卖家取消订单
    switch ($row['status']){
        case 0:
            $row['status'] = '待确认';
            break;
        case 1:
            $row['status'] = '已确认/待付款';
            break;
        case 2:
            $row['status'] = '已付款/待发货';
            break;
        case 3:
            $row['status'] = '发货中';
            break;
        case 4:
            $row['status'] = '已发货';
            break;
        case 5:
            $row['status'] = '买家确认收货';
            break;
        case 6:
            $row['status'] = '订单完成';
            break;
        case 7:
            $row['status'] = '买家取消订单';
            break;
        case 8:
            $row['status'] = '卖家取消订单';
            break;
        default:
            $row['status'] = '其他未知错误';
            break;
    }
    $data[$i]['status'] = $row['status'];
    $data[$i]['rev_name'] = $row['rev_name'];

    $data[$i]['rev_addr'] = $row['rev_addr'];
    $data[$i]['rev_mail'] = $row['rev_mail'];
    $data[$i]['rev_post'] = $row['rev_post'];
    $data[$i]['rev_mobile'] = $row['rev_mobile'];
    $data[$i]['account_name'] = $row['account_name'];
    $data[$i]['brands_name'] = $row['brands_name'];
    $data[$i]['project'] = $row['project'];
    $i++;
}
/*echo "<pre>";
print_r($data);
echo "</pre>";*/

//设置excel列名
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1','订单ID');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('B1','订单编号');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('C1','状态');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('D1','收货人');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('E1','收货地址');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('F1','收货人邮箱');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('G1','邮编');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('H1','收货人电话');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('I1','会员帐号');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('J1','品牌id');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('K1','项目名');
//背景填充颜色
$objPHPExcel->getActiveSheet()->getStyle( 'A1:K1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objPHPExcel->getActiveSheet()->getStyle( 'A1:K1')->getFill()->getStartColor()->setARGB('FF808080');
//把数据循环写入excel中
foreach($data as $key => $value){
    $key+= 2;   //从第二行开始填充
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$key,$value['id']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B'.$key,$value['order_sn']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('C'.$key,$value['status']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('D'.$key,$value['rev_name']);

    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('E'.$key,$value['rev_addr']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('F'.$key,$value['rev_mail']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('G'.$key,$value['rev_post']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('H'.$key,$value['rev_mobile']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('I'.$key,$value['account_name']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('J'.$key,$value['brands_name']);
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('k'.$key,$value['project']);
}
//设置默认字体
$objPHPExcel->getDefaultStyle()->getFont()->setName( 'Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(12);

//设置列宽
$objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
//设置居中
$objPHPExcel->getDefaultStyle()->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
//excel保存在根目录下  如要导出文件，以下改为注释代码
//$objPHPExcel->getActiveSheet() -> setTitle('SetExcelName');
//$objPHPExcel-> setActiveSheetIndex(0);
//$objWriter = $iofactory -> createWriter($objPHPExcel, 'Excel2007');
//$objWriter -> save('SetExcelName.xlsx');
//导出代码
$objPHPExcel->getActiveSheet() -> setTitle('订单列表');
$objPHPExcel-> setActiveSheetIndex(0);

$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
$filename = '订单列表.xlsx';
ob_end_clean();//清除缓存以免乱码出现
header('Content-Type: application/vnd.ms-excel');
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Cache-Control: max-age=0');
$objWriter -> save('php://output');
?>