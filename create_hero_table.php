<?
require 'vendor/autoload.php';

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("FizMiG_v_2_3.xlsx");

$worksheet = $spreadsheet->getSheetByName('1.8');

$data = $worksheet->toArray();

/*    */
$data = array_slice($data, 7);

$keys = [];

for ($i=0; $i < count($worksheet->getDrawingCollection()); $i++) {     
    $keys[] = $worksheet->getDrawingCollection()[$i]->getCoordinates();
}

natsort($keys);

$heroes = [];
$num = 0;

foreach($keys as $key => $value){

    $hero = [];
    $hero['name'] = $data[$num][2];
    $hero['class']  =$data[$num][3];
    $hero['sex'] = $data[$num][4];
    $hero['race']  =$data[$num][5];
    $hero['bio']  =$data[$num][6];

    $drawing = $worksheet->getDrawingCollection()[$key];      
        
    
    $zipReader = fopen($drawing->getPath(), 'r');
    $imageContents = '';
    while (!feof($zipReader)) {
        $imageContents .= fread($zipReader, 1024);
    }
    fclose($zipReader);
    $extension = $drawing->getExtension();

    $u= "data:image/jpeg;base64," . base64_encode($imageContents);
    
    $hero['img_data'] = $u;
    $heroes[] = $hero;
    $num++;
}

try{

    $conn = new PDO("mysql:host=$host;dbname=$db", $login, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $sql = "CREATE TABLE heroes (
        id INT NOT NULL AUTO_INCREMENT,
        name VARCHAR(40),
        class TEXT,
        sex TEXT,
        race TEXT,
        bio TEXT,
        img_data TEXT,
        PRIMARY KEY (`id`)
    )";
   $conn->exec($sql);

} catch(PDOException $e) {
  echo $sql . "<br>" . $e->getMessage();
}

foreach($heroes as $key => $value){
    
    $sql = "INSERT INTO heroes(name, class, sex, race, bio, img_data)
                        VALUES(:name, :class, :sex, :race, :bio, :img_data)";
    $sql = $conn->prepare($sql);
    $sql->execute(array(
        ":name" => $value['name'],
        ":class" => $value['class'],
        ":sex" => $value['sex'],
        ":race" => $value['race'],
        ":bio" => $value['bio'],
        ":img_data" => $value['img_data'],
    ));
}
      