<?php
	ini_set("display_errors",true);
	class open_conn{
		protected $conn;
		protected $host;
		protected $port;
		protected $db;
		protected $user;
		protected $pass;
		
		function __construct($dbtype=""){
			$this->get_connection_desc($dbtype);
		}
		
		function get_conn(){
			$this->conn=$this->get_db_connection();
			return $this->conn;
		}
		
		function __destruct(){
			pg_close($this->conn);
		}
		
		function get_connection_desc($db_type=""){
			if ($db_type=="sms"){
				$this->host="192.168.1.1";
				$this->port=5432;
				$this->user="crm";
				$this->db="smsd";
				$this->pass="Ys70A0u3FyIvV4BewH1X";
			}else{
				$this->host="192.168.1.1";
				$this->port=5432;
				$this->user="crm";
				$this->db="crm";
				$this->pass="Ys70A0u3FyIvV4BewH1X";			
			}		
		}
		
		function get_db_connection(){
			$conn=pg_connect("host=".$this->host." port=".$this->port." user=".$this->user." password=".$this->pass." dbname=".$this->db);
			if ($conn){
				return $conn;
			}else{
				echo "DATABASE IS NOT CONNECTED!!";
				exit;
			}
		}
	}
?>