<?
require_once('connection_data.php');

try{
    $conn = new PDO("mysql:host=$host", $login, $password);
    $conn->exec("CREATE DATABASE $db;");
}
catch (PDOException $e) {
    die("DB ERROR: ". $e->getMessage());
}


