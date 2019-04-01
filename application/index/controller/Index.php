<?php
namespace app\index\controller;
use think\Controller;
use think\Loader;

use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Cell;
use PHPExcel_Writer_Excel5;
use PHPExcel_Writer_Excel2007;

class Index
{


    public function index()
    {


      if(request()->isPost()) {
           
            Loader::import('PHPExcel.PHPExcel');
            Loader::import('PHPExcel.PHPExcel.PHPExcel_IOFactory');
            Loader::import('PHPExcel.PHPExcel.PHPExcel_Cell');
            //实例化PHPExcel
            $objPHPExcel = new \PHPExcel();
            $file = request()->file('excel');

            if ($file) {
                
                $file_types = explode(".", $_FILES ['excel'] ['name']); // ["name"] => string(25) "excel文件名.xls"
                $file_type = $file_types [count($file_types) - 1];//xls后缀
                $file_name = $file_types [count($file_types) - 2];//xls去后缀的文件名
                /*判别是不是.xls文件，判别是不是excel文件*/
                if (strtolower($file_type) != "xls" && strtolower($file_type) != "xlsx") {
                    echo '不是Excel文件，重新上传';
                    die;
                }

                $info = $file->move(ROOT_PATH . 'public' . DS . 'excel');//上传位置
                $path = ROOT_PATH . 'public' . DS . 'excel' . DS;
                $file_path = $path . $info->getSaveName();//上传后的EXCEL路径
                //echo $file_path;//文件路径

                //获取上传的excel表格的数据，形成数组
                $re = $this->actionRead($file_path, 'utf-8');
                //dump($re);exit;
                array_splice($re, 1, 0);
                unset($re[0]);
				//dump($re); //查看数组
				
                /*将数组的键改为自定义名称*/
                $keys = array('username', 'zhiye', 'address');
                foreach ($re as $i => $vals) {
                    $re[$i] = array_combine($keys, $vals);
                }
				//遍历数组写入数据库
                for ($i = 1; $i < count($re); $i++) {
                    $data = $re[$i];
                    $res = db('username')->insert($data);
              
                }

            }
        }

       return view();
       //return $this->fetch();
    }



public function actionRead($filename, $encode = 'utf-8')
    {
        $objReader = PHPExcel_IOFactory::createReader('Excel2007');
        $objReader->setReadDataOnly(true);
        $objPHPExcel = $objReader->load($filename);
        $objWorksheet = $objPHPExcel->getActiveSheet();
        $highestRow = $objWorksheet->getHighestRow();
         $highestColumn = $objWorksheet->getHighestColumn();
         $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);
         $excelData = array();
         for($row = 1; $row <= $highestRow; $row++)
         {
         for ($col = 0; $col < $highestColumnIndex; $col++)
         {
         $excelData[$row][]=(string)$objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
         }
         }
         return $excelData;
    }


///index.php/index/index/out
//导出
public function out(){
          $path = dirname(__FILE__); //找到当前脚本所在路径
          Loader::import("PHPExcel.PHPExcel");
          Loader::import("PHPExcel.PHPExcel.Writer.IWriter");
          Loader::import("PHPExcel.PHPExcel.Writer.Abstract");
          Loader::import("PHPExcel.PHPExcel.Writer.Excel5");
          Loader::import("PHPExcel.PHPExcel.Writer.Excel2007");
          Loader::import("PHPExcel.PHPExcel.IOFactory");
          $objPHPExcel = new \PHPExcel();
          
          $objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
          
          $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);

          // 实例化完了之后就先把数据库里面的数据查出来
          $sql = db('username')->select();
               
          // 设置表头信息
          $objPHPExcel->setActiveSheetIndex(0)
              ->setCellValue('A1', 'id')
              ->setCellValue('B1', '姓名')
              ->setCellValue('C1', '职业')
              ->setCellValue('D1', '地址');
         /*     ->setCellValue('E1', 'salt')
              ->setCellValue('F1', 'avatar')
              ->setCellValue('G1', 'email')
              ->setCellValue('H1', 'loginfailure')
              ->setCellValue('I1', 'logintime')
              ->setCellValue('J1', 'createtime')
              ->setCellValue('K1', 'updatetime')
              ->setCellValue('L1', 'token')
              ->setCellValue('M1', 'status');*/
          $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
          $count = count($sql);  //计算有多少条数据
          for ($i = 2; $i <= $count+1; $i++) {
              $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $sql[$i-2]['id']);
              $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $sql[$i-2]['username']);
              $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $sql[$i-2]['zhiye']);
              $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $sql[$i-2]['address']);
   /*           $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $sql[$i-2]['salt']);
              $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $sql[$i-2]['avatar']);
              $objPHPExcel->getActiveSheet()->setCellValue('G' . $i, $sql[$i-2]['email']);
              $objPHPExcel->getActiveSheet()->setCellValue('H' . $i, $sql[$i-2]['loginfailure']);
              $objPHPExcel->getActiveSheet()->setCellValue('I' . $i, $sql[$i-2]['logintime']);
              $objPHPExcel->getActiveSheet()->setCellValue('J' . $i, $sql[$i-2]['createtime']);
              $objPHPExcel->getActiveSheet()->setCellValue('K' . $i, $sql[$i-2]['updatetime']);
              $objPHPExcel->getActiveSheet()->setCellValue('L' . $i, $sql[$i-2]['token']);
              $objPHPExcel->getActiveSheet()->setCellValue('M' . $i, $sql[$i-2]['status']);*/
          }
          $objPHPExcel->getActiveSheet()->setTitle('username');      //设置sheet的名称
          $objPHPExcel->setActiveSheetIndex(0);                   //设置sheet的起始位置
          $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   //通过PHPExcel_IOFactory的写函数将上面数据写出来
          $PHPWriter = PHPExcel_IOFactory::createWriter( $objPHPExcel,"Excel2007");
          header('Content-Disposition: attachment;filename="admin.xlsx"');
          header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          $res=$PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
          if($res){
           echo '导出成功';
          }else{
          	echo '导出失败';
          }
          exit;
      }










}
