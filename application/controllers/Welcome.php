<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Welcome extends CI_Controller {


	public function index()
	{
		$this->load->view('welcome_message');
	}

	public function geraPlanilha(){
    	##e#cho BASEPATH; DIE;
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
			$planilha->setActiveSheetIndex(0)->setCellValue( 'A' . $c, $linha['email'] );
		endforeach;

		$planilha->getActiveSheet()->setTitle('Planilha 1');

		$objGravar = PHPExcel_IOFactory::createWriter($planilha, 'Excel2007');
		$objGravar->save($arquivo);
		
		//echo 'planilha Gerada com sucesso !  :)';

	}
}
