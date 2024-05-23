public function exportCustomers(){

		$customers_model = new CustomersModel; 	

		$getResult = $customers_model->getCustomersRecord();

		$spreadsheet = new Spreadsheet();

      	$objPHPExcel = $spreadsheet->getActiveSheet();
        // $objPHPExcel->setActiveSheetIndex(0);           
        
        $objPHPExcel->getStyle('A1')->applyFromArray
        (
            array ('font' => array('size' => 11,'bold' => true,'color' => array('rgb' => '000000')))
        );//report orders styling

        $styleArray2 = array(                   
            'font'  => array(
                'bold'  => true,
                'size'  => 14
            )
        );

        $styleArray3 = array(                   
            'font'  => array(
                'bold'  => true,
                'size'  => 11
            )
        );
        
        // font bold & line height big
        $variable1 = array(1);
        foreach ($variable1 as $key1 => $value1) {
            $objPHPExcel->getRowDimension($value1)->setRowHeight(20);
            $objPHPExcel->getStyle('A'.$value1.':H'.$value1)->applyFromArray($styleArray2);    
        }


        $objPHPExcel->SetCellValue('A1', "Full Name");
        $objPHPExcel->SetCellValue('B1', "Email");
        $objPHPExcel->SetCellValue('C1', "BirthDate(yy-mm-dd)");
        $objPHPExcel->SetCellValue('D1', "Billing Address");
        $objPHPExcel->SetCellValue('E1', "Shipping Address");
        $objPHPExcel->SetCellValue('F1', "Last Order");
        $objPHPExcel->SetCellValue('G1', "Total Spent");
        $objPHPExcel->SetCellValue('H1', "Average order Value");

     	$rowId = 2;
        foreach ($getResult as $row) {
        	//$fullName = implode($row['firstName']," ",$row['lastName']);
        	$fullName = $row['firstName'] . ' ' . $row['lastName'];
            $objPHPExcel->SetCellValue('A'.$rowId, $fullName);
            $objPHPExcel->SetCellValue('B'.$rowId, $row['email']);
            $objPHPExcel->SetCellValue('C'.$rowId, $row['dateadd']);
            $objPHPExcel->SetCellValue('D'.$rowId, $row['billadd']);
            $objPHPExcel->SetCellValue('E'.$rowId, $row['shipadd']);
            $objPHPExcel->SetCellValue('F'.$rowId, $row['lsorder']);
            $objPHPExcel->SetCellValue('G'.$rowId, $row['totalsp']);
            $objPHPExcel->SetCellValue('H'.$rowId, $row['avgorder']);

         //    if(!empty($row['productImage'])){
	        //     $objPHPExcel->SetCellValue('E' . $rowId,base_url('uploads/products').'/'.$row['productImage']);
	        //     $objPHPExcel->getCell('E' . $rowId)->getHyperlink()->setUrl(base_url('uploads/products').'/'.$row['productImage']);
	        // }

            $rowId++;
        }

        $objPHPExcel->getColumnDimension('A')->setWidth(30);
        $objPHPExcel->getColumnDimension('B')->setWidth(30);
        $objPHPExcel->getColumnDimension('C')->setWidth(30);
        $objPHPExcel->getColumnDimension('D')->setWidth(30);
        $objPHPExcel->getColumnDimension('E')->setWidth(30);
        $objPHPExcel->getColumnDimension('F')->setWidth(30);
        $objPHPExcel->getColumnDimension('G')->setWidth(30);
        $objPHPExcel->getColumnDimension('H')->setWidth(30);

        $objWriter = new Xlsx($spreadsheet);           
        
        $fileName = "Export-All-customers.xlsx";
        
        header("Content-Type: application/vnd.ms-excel");
        header('Content-Disposition: attachment;filename="'.$fileName.'"');
        header('Cache-Control: max-age=0');
      
        $objWriter->save('uploads/product_report/' . $fileName);
       
        $xls_url = '/uploads/product_report/' . $fileName;
        echo $xls_url;

	}
