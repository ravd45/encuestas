<?php
class Encuesta_model extends CI_Model
{
    function __construct(){
        parent::__construct();
    }

    function get_encuesta($id_encuesta){
      $str_query = "SELECT * FROM pregunta WHERE id_encuesta = ? order by orden";
      return $this->db->query($str_query, array($id_encuesta))->result_array();
    }// get_encuesta()

    function acceso_sistema($id_cct){
    	$status = 1;
    	$fcreacion = date("Y-m-d H:i:s");

    	$str_query = "SELECT * FROM sin_registro WHERE id_cct = ?";
    	$registro = $this->db->query($str_query, array($id_cct))->result_array();

    	if(count($registro) > 0){
    		$str_query2 = "UPDATE sin_registro SET registro = ?, fec_ultimoreg = ? WHERE id_cct = ?";
    		return $this->db->query($str_query2, array($status, $fcreacion, $id_cct));
    	}else{	
    		$str_query2 = "INSERT INTO sin_registro(id_cct, registro, fec_ultimoreg) VALUES(?, ?, ?)";
    		return $this->db->query($str_query2, array($id_cct, $status, $fcreacion));
    	}
    }

    function set_encuesta($idcct, $cct, $respuestas, $idencuesta){
    	$this->db->trans_start();
    	$str_query = "INSERT INTO encuesta_x_cct (id_cct, id_encuesta) VALUES(?, ?);";
		$this->db->query($str_query, array($idcct, $idencuesta));
		$ultimoId = $this->db->insert_id();
		foreach ($respuestas as $respuesta) {
			$str_query = "INSERT INTO respuesta (respuesta, id_pregunta, id_aplica) VALUES('{$respuesta['respuesta']}', {$respuesta['id_pregunta']}, {$ultimoId});";
			$this->db->query($str_query);
		}
		
		$this->db->trans_complete();

	    if ($this->db->trans_status() === FALSE)
	    {
	        return false;
	    }else{
	        return true;
	    }
    }

    function get_encuestasxcct($idcct){
    	 $str_query = "SELECT  e.id_aplica AS id_encuesta, GROUP_CONCAT(CONCAT(r.id_pregunta, '_', r.respuesta) SEPARATOR '&') AS resp FROM encuesta_x_cct e
		INNER JOIN respuesta r ON r.id_aplica = e.id_aplica
		WHERE e.id_cct = ?
		GROUP BY e.id_aplica";
      return $this->db->query($str_query, array($idcct))->result_array();
    }

    function delete_encuestaxcct($id_encuesta){
    
		$this->db->trans_start();
    	$str_query = "DELETE FROM respuesta WHERE id_aplica = ?";
		$this->db->query($str_query, array($id_encuesta));

		$str_query = "DELETE FROM encuesta_x_cct WHERE id_aplica = ?";
		$this->db->query($str_query, array($id_encuesta));
		
		$this->db->trans_complete();

	    if ($this->db->trans_status() === FALSE)
	    {
	        return false;
	    }else{
	        return true;
	    }

    }

    function get_encuestaxcct($id_encuesta){
    	$str_query = "SELECT * FROM respuesta WHERE id_aplica = ?";
    	return $this->db->query($str_query, array($id_encuesta))->result_array();
    }

    function set_encuesta_edit($respuestas, $id_editando){
    	$this->db->trans_start();
			foreach ($respuestas as $respuesta) {
				$str_query = "UPDATE respuesta SET respuesta = '{$respuesta['respuesta']}' WHERE id_pregunta = {$respuesta['id_pregunta']} AND id_aplica ={$id_editando}";
				$this->db->query($str_query);
			}
		$this->db->trans_complete();

	    if ($this->db->trans_status() === FALSE)
	    {
	        return false;
	    }else{
	        return true;
	    }
    	
    }

    function sin_registros_update($id_cct, $status){
    	$fcreacion = date("Y-m-d H:i:s");

    	$str_query = "SELECT * FROM sin_registro WHERE id_cct = ?";
    	$registro = $this->db->query($str_query, array($id_cct))->result_array();

    	if(count($registro) > 0){
    		$str_query2 = "UPDATE sin_registro SET registro = ?, fec_ultimoreg = ? WHERE id_cct = ?";
    		return $this->db->query($str_query2, array($status, $fcreacion, $id_cct));
    	}else{	
    		$str_query2 = "INSERT INTO sin_registro(id_cct, registro, fec_ultimoreg) VALUES(?, ?, ?)";
    		return $this->db->query($str_query2, array($id_cct, $status, $fcreacion));
    	}
    }

    function get_reporte_excel(){
    	$str_query = "SELECT x.*, m.municipio AS nom_municipio
						FROM(
						SELECT
						ct.id_cct, sr.fec_ultimoreg, en.id_aplica, ct.cct, ct.turno, ct.nombre_ct, 
						MAX(CASE WHEN p.id_pregunta = 18 THEN r.respuesta ELSE '' END) AS 'NOMBRE(S)',
						MAX(CASE WHEN p.id_pregunta = 19 THEN r.respuesta ELSE '' END) AS 'PRIMER APELLIDO',
						MAX(CASE WHEN p.id_pregunta = 20 THEN r.respuesta ELSE '' END) AS 'SEGUNDO APELLIDO',
						MAX(CASE WHEN p.id_pregunta = 21 THEN r.respuesta ELSE '' END) AS 'EDAD(más de 15 años)',
						MAX(CASE WHEN p.id_pregunta = 22 THEN r.respuesta ELSE '' END) AS 'DOMICILIO(calle y número)',
						MAX(CASE WHEN p.id_pregunta = 23 THEN r.respuesta ELSE '' END) AS 'COLONIA',
						MAX(CASE WHEN p.id_pregunta = 24 THEN r.respuesta ELSE '' END) AS 'MUNICIPIO',
						MAX(CASE WHEN p.id_pregunta = 25 THEN r.respuesta ELSE '' END) AS 'LOCALIDAD',
						MAX(CASE WHEN p.id_pregunta = 26 THEN r.respuesta ELSE '' END) AS 'TELÉFONO',
						MAX(CASE WHEN p.id_pregunta = 27 THEN r.respuesta ELSE '' END) AS 'REZAGO'
						FROM encuesta_x_cct en
						INNER JOIN cct ct ON en.id_cct=ct.id_cct
						INNER JOIN respuesta r ON en.id_aplica= r.id_aplica
						INNER JOIN pregunta p ON r.id_pregunta = p.id_pregunta
						LEFT JOIN sin_registro sr on en.id_cct = sr.id_cct
						GROUP BY r.id_aplica
						ORDER BY en.id_aplica, ct.cct, ct.id_turno) AS x
						INNER JOIN municipio m ON x.MUNICIPIO=m.id_municipio";
		return $this->db->query($str_query)->result_array();
    }

}// Prioridad_model
