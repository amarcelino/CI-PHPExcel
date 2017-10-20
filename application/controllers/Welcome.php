<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Welcome extends CI_Controller {


	public function index()
	{
		$this->load->view('welcome_message');
	}

	public function geraPlanilha(){
		$this->load->library('PHPExcel');
		$arquivo = './uploads/relatorio.xlsx';
		$planilha = $this->phpexcel;

		$registros = [
			['nome'=>'André Marcelino', 'email'=>'marcelino@gmail.com'],
			['nome'=>'Nome Teste 1', 'email'=>'t1@gmail.com'],
			['nome'=>'Nome Teste 2', 'email'=>'t2@gmail.com']
		];

		$planilha->setActiveSheetIndex(0)->setCellValue('A1', 'Nome');
		$planilha->setActiveSheetIndex(0)->setCellValue('B1', 'Email');

		$c = 1;
		foreach( $registros as $linha ):
			$c++;
			$planilha->setActiveSheetIndex(0)->setCellValue( 'A' . $c, $linha['nome'] );
			$planilha->setActiveSheetIndex(0)->setCellValue( 'B' . $c, $linha['email'] );
		endforeach;

		$planilha->getActiveSheet()->setTitle('Planilha 1');

		$objGravar = PHPExcel_IOFactory::createWriter($planilha, 'Excel2007');
		$objGravar->save($arquivo);
		
		//echo 'planilha Gerada com sucesso !  :)';

	}
    
	
	
	public function excelArray( $param ) {
        // Here is the sample array of data
        $data = array(
            array( 'name' => 'A', 'mail' => 'a@gmail.com', 'age' => 43 ),
            array( 'name' => 'C', 'mail' => 'c@gmail.com', 'age' => 24 ),
            array( 'name' => 'B', 'mail' => 'b@gmail.com', 'age' => 35 ),
            array( 'name' => 'G', 'mail' => 'f@gmail.com', 'age' => 22 ),
            array( 'name' => 'F', 'mail' => 'd@gmail.com', 'age' => 52 ),
            array( 'name' => 'D', 'mail' => 'g@gmail.com', 'age' => 32 ),
            array( 'name' => 'E', 'mail' => 'e@gmail.com', 'age' => 34 ),
            array( 'name' => 'K', 'mail' => 'j@gmail.com', 'age' => 18 ),
            array( 'name' => 'L', 'mail' => 'h@gmail.com', 'age' => 25 ),
            array( 'name' => 'H', 'mail' => 'i@gmail.com', 'age' => 28 ),
            array( 'name' => 'J', 'mail' => 'j@gmail.com', 'age' => 53 ),
            array( 'name' => 'I', 'mail' => 'l@gmail.com', 'age' => 26 ),
        );


        $objPHPExcel = new PHPExcel();
        // Fill worksheet from values in array
        $objPHPExcel->getActiveSheet()->fromArray( $data, null, 'A2' );

        // Rename worksheet
        $objPHPExcel->getActiveSheet()->setTitle( 'enois' );

        // Set AutoSize for name and email fields
        $objPHPExcel->getActiveSheet()->getColumnDimension( 'A' )->setAutoSize( true );
        $objPHPExcel->getActiveSheet()->getColumnDimension( 'B' )->setAutoSize( true );


        // Save Excel 2007 file
        $objWriter = PHPExcel_IOFactory::createWriter( $objPHPExcel, 'Excel2007' );
        $objWriter->save( 'MyExcel.xls' );
    }	
	
}
