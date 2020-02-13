<?php
defined('BASEPATH') or exit('No direct script access allowed');

class Reportes extends CI_Controller
{

    public function __construct()
    {
        parent::__construct();
        $this->load->library('Utilerias');
        $this->load->library('My_PHPExcel');
        $this->load->model('Encuesta_model');
        $this->cct = array();
        $this->sesion = Utilerias::get_cct_sesion($this);
    } // __construct()

    function get_reporte_excel() {
    if (Utilerias::verifica_sesion_redirige($this)) {
                    $this->styleArray_encabezado = array(
                        'borders' => array(
                                'allborders' => array(
                                        'style' => PHPExcel_Style_Border::BORDER_THIN
                                )
                        ),
                        'fill' => array(
                            'type'  => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array(
                                        'rgb' => 'F4FCFC')
                    ),
                    'font' => array(
                            'name'  => 'Arial',
                            'bold'  => true,
                            'color' => array(
                                    'rgb' => '000000'
                            )
                    ),
                    'alignment' =>  array(
                            'horizontal'=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                            'vertical'  => PHPExcel_Style_Alignment::VERTICAL_CENTER
                    )
                );
                $this->styleArray_contenido = array(
                    'borders' => array(
                            'allborders' => array(
                                    'style' => PHPExcel_Style_Border::BORDER_THIN
                            )
                    ),
                    'fill' => array(
                        'type'  => PHPExcel_Style_Fill::FILL_SOLID
                ),
                'font' => array(
                        'name'  => 'Arial',
                        'color' => array(
                                'rgb' => '000000'
                        )
                )
            );

            $this->styleArray_titulo = array(
                'borders' => array(
                        'allborders' => array(
                                'style' => PHPExcel_Style_Border::BORDER_THIN
                        )
                ),
                'fill' => array(
                    'type'  => PHPExcel_Style_Fill::FILL_SOLID,
                    'color' => array(
                                'rgb' => 'DFF5F5')
            ),
            'font' => array(
                    'name'  => 'Arial',
                    'bold'  => true,
                    'color' => array(
                            'rgb' => '000000'
                    )
            ),
            'alignment' =>  array(
                    'horizontal'=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    'vertical'  => PHPExcel_Style_Alignment::VERTICAL_CENTER
            )
            );

            $obj_phpexcel = new My_PHPExcel();
            $obj_phpexcel->getActiveSheet()->SetCellValue('A1', 'REPORTE DE CAPTURA DE ENCUESTAS');
            $obj_phpexcel->getActiveSheet()->SetCellValue('A2', 'ID CCT');
            $obj_phpexcel->getActiveSheet()->SetCellValue('B2', 'FECHA DE REGISTRO');
            $obj_phpexcel->getActiveSheet()->SetCellValue('C2', 'CCT');
            $obj_phpexcel->getActiveSheet()->SetCellValue('D2', 'Turno');
            $obj_phpexcel->getActiveSheet()->SetCellValue('E2', 'Nombre cct');
            $obj_phpexcel->getActiveSheet()->SetCellValue('F2', 'Nombre');
            $obj_phpexcel->getActiveSheet()->SetCellValue('G2', 'Apellido1');
            $obj_phpexcel->getActiveSheet()->SetCellValue('H2', 'Apellido2');
            $obj_phpexcel->getActiveSheet()->SetCellValue('I2', 'Edad');

            $obj_phpexcel->getActiveSheet()->SetCellValue('J2', 'Domicilio');
            $obj_phpexcel->getActiveSheet()->SetCellValue('K2', 'Municipio');
            $obj_phpexcel->getActiveSheet()->SetCellValue('L2', 'Colonia');
            $obj_phpexcel->getActiveSheet()->SetCellValue('M2', 'Localidad');
            $obj_phpexcel->getActiveSheet()->SetCellValue('N2', 'Telefono');
            $obj_phpexcel->getActiveSheet()->SetCellValue('O2', 'Rezago');
            $arr_datos = $this->Encuesta_model->get_reporte_excel();
            $indice = 3;
            foreach ($arr_datos as $item) {
                    $obj_phpexcel->getActiveSheet()->SetCellValue('A'.$indice, ($item['id_cct']));    
                    $obj_phpexcel->getActiveSheet()->SetCellValue('B'.$indice, ($item['fec_ultimoreg']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('C'.$indice, ($item['cct']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('D'.$indice, "  ".($item['turno']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('E'.$indice, "  ".($item['nombre_ct']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('F'.$indice, "  ".($item['NOMBRE(S)']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('G'.$indice, "  ".($item['PRIMER APELLIDO']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('H'.$indice, "  ".($item['SEGUNDO APELLIDO']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('I'.$indice, "  ".($item['EDAD(más de 15 años)']));

                    $obj_phpexcel->getActiveSheet()->SetCellValue('J'.$indice, "  ".($item['DOMICILIO(calle y número)']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('K'.$indice, "  ".($item['nom_municipio']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('L'.$indice, "  ".($item['COLONIA']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('M'.$indice, "  ".($item['LOCALIDAD']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('N'.$indice, "  ".($item['TELÉFONO']));
                    $obj_phpexcel->getActiveSheet()->SetCellValue('O'.$indice, "  ".($item['REZAGO']));
                    $indice++;
            }


            $obj_phpexcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);

            $obj_phpexcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
            $obj_phpexcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);

            $fcreacion = date("Y-m-d_H-i-s");
            $nombre_excel = "reporte_encuestas_{$fcreacion}.xlsx";

            $obj_phpexcel->getActiveSheet()->mergeCells('A1:O1');
            $obj_phpexcel->getActiveSheet()->getStyle('A1:O1')->applyFromArray($this->styleArray_titulo);
            $obj_phpexcel->getActiveSheet()->getStyle('A2:O2')->applyFromArray($this->styleArray_encabezado);
            $obj_phpexcel->getActiveSheet()->getStyle('A3:O'.($indice-1))->applyFromArray($this->styleArray_contenido);
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename='.$nombre_excel);
            header('Cache-Control: max-age=0');
            $obj_writer=PHPExcel_IOFactory::createWriter($obj_phpexcel,'Excel2007');
            $obj_writer->save('php://output');
            exit;

    }// verifica_sesion_redirige
}// get_reporte_excel()




} // class
