<?

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("FizMiG_v_2_3.xlsx");

$worksheet = $spreadsheet->getSheetByName('1.6.1');

$data = $worksheet->toArray();

function clear_data($data){
    $result = [];
    foreach($data as $key => $value){
        if($key > 4 && $value[1] != NULL){
            $result[] = $value;
        }
    }
    return $result;
}

$data = clear_data($data);

function parse_item($arr, $num ,$worksheet){
    $image_num;
    $item = [];
    foreach($arr as $key => $value){
        switch($key){
            case 1:
                $item['name']  = $arr[$key];
                break;
            case 3:
                $item['without']  = $arr[$key];
                break;
            case 4:
                $item['base']  = $arr[$key];
                break;
            case 5:
                $item['advanced']  = $arr[$key];
                break;
            case 6:
                $item['expert']  = $arr[$key];
                break;
        }
        switch($num){   
            case 0:
                $image_num = 20;
                break;
            case 1:
                $image_num = 0;
                break;
            case 2:
                $image_num = 23;
                break;
            case 3:
                $image_num = 17;
                break;                
            case 4:
                $image_num = 1;
                break;
            case 5:
                $image_num = 22;
                break;
            case 6:
                $image_num = 2;
                break;
            case 7:
                $image_num = 3;
                break;
            case 8:
                $image_num = 27;
                break;                
            case 9:
                $image_num = 4;
                break;
            case 10:
                $image_num = 5;
                break;
            case 11:
                $image_num = 15;
                break;
            case 12:
                $image_num = 14;
                break;
            case 13:
                $image_num = 16;
                break;                
            case 14:
                $image_num = 13;
                break;
            case 15:
                $image_num = 7;
                break;
            case 16:
                $image_num = 12;
                break;
            case 17:
                $image_num = 8;
                break;
            case 18:
                $image_num = 21;
                break;                
            case 19:
                $image_num = 26;
                break;
            case 20:
                $image_num = 25;
                break;                
            case 21:
                $image_num = 10;
            break;
            case 22:
                $image_num = 11;
            break;
            case 23:
                $image_num = 24;
            break;
            case 24:
                $image_num = 19;
                break;
            case 25:
                $image_num = 18;
            break;                
            case 26:
                $image_num = 6;
            break;
            case 27:
                $image_num = 9;
            break;
            
        }
        $drawing = $worksheet->getDrawingCollection()[$image_num];

        $zipReader = fopen($drawing->getPath(), 'r');
        $imageContents = '';

        while (!feof($zipReader)) {
        $imageContents .= fread($zipReader, 1024);
        }
        fclose($zipReader);
        $extension = $drawing->getExtension();
    
        $item['img_data'] = "data:image/jpeg;base64," . base64_encode($imageContents);
    }
    return $item;  
}

function parse_all($arr, $worksheet){

    $result = [];

    foreach($arr as $key => $value){
        $result[] = parse_item($value, $key, $worksheet);
    }
    return $result;
}

$result = parse_all($data, $worksheet);

try{

    $conn = new PDO("mysql:host=$host;dbname=$db", $login, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $sql = "CREATE TABLE second_skills (
        id INT NOT NULL AUTO_INCREMENT,
        name VARCHAR(40),
        without TEXT,
        base TEXT,
        advanced TEXT,
        expert TEXT,
        img_data TEXT,
        PRIMARY KEY (`id`)
    )";
   $conn->exec($sql);

} catch(PDOException $e) {
  echo $sql . "<br>" . $e->getMessage();
}

foreach($result as $key => $value){
    
    $sql = "INSERT INTO second_skills(name, without, base, advanced, expert, img_data)
                        VALUES(:name, :without, :base, :advanced, :expert, :img_data)";
    $sql = $conn->prepare($sql);
    $sql->execute(array(
        ":name" => $value['name'],
        ":without" => $value['without'],
        ":base" => $value['base'],
        ":expert" => $value['expert'],
        ":advanced" => $value['advanced'],
        ":img_data" => $value['img_data'],
    ));
}