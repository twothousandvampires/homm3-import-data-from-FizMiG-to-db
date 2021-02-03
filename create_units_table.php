<?

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("FizMiG_v_2_3.xlsx");

$worksheet = $spreadsheet->getSheetByName('3.1');

$data = $worksheet->toArray();

function parseUnit($item, $num, $worksheet){
    $unit = [];
    for ($i=0; $i < count($item); $i++) {       
        switch($i){
            case 1:
                $unit['name'] = $item[$i];
                break;
            case 3:
                $unit['level'] = $item[$i];
                break;
            case 4:
                $unit['fraction'] = $item[$i];
                break;
            case 5:
                $unit['min_damage'] = $item[$i];
                break;
            case 6:
                $unit['max_damage'] = $item[$i];
                break;
            case 7:
                $unit['attack'] = $item[$i];
                break;
            case 8:
                $unit['defence'] = $item[$i];
                break;
            case 9:
                $unit['hp'] = (int)$item[$i];
                break;
            case 10:
                $unit['speed'] = $item[$i];
                break;
            case 11:
                $unit['growth'] = $item[$i];
                break;
            case 12:
                    $unit['gold'] = $item[$i];
                break;
            case 13:
                    $unit['resources'] = $item[$i];
                break;
            case 14:
                    $unit['ability'] = $item[$i];
                break;
            case 15:
                    $unit['ai_value'] = $item[$i];
                break;
            
        }
        $drawing = $worksheet->getDrawingCollection()[$num];

        $zipReader = fopen($drawing->getPath(), 'r');
        $imageContents = '';
        while (!feof($zipReader)) {
            $imageContents .= fread($zipReader, 1024);
        }
        fclose($zipReader);
        $extension = $drawing->getExtension();


        $unit['img_data'] = "data:image/jpeg;base64," . base64_encode($imageContents);
       
    }
    return $unit;
}

function getUnitsArray($arr){
    $unitsArray = [];
    $start;
    $end;
    foreach ($arr as $key=>$value) {
        if($value[1] === 'Копейщик') $start = $key;
        if($value[1] === 'Лазурный дракон') $end = $key;
    }
    return array_slice(array_slice($arr, $start), 0 , $end - $start +1);
}

function parseUnitsArray($data, $worksheet){

    $array_units = getUnitsArray($data);
    $units = [];

    foreach ($array_units as $key => $value) {
        $units[] = parseUnit($value, $key, $worksheet);
    }

    return $units;
}

$result = parseUnitsArray($data, $worksheet);

try{

    $conn = new PDO("mysql:host=$host;dbname=$db", $login, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $sql = "CREATE TABLE units (
        id INT NOT NULL AUTO_INCREMENT,
        name VARCHAR(40),
        level INT,
        fraction VARCHAR(40),
        min_damage INT,
        max_damage INT,
        attack INT,
        defence INT,
        hp INT,
        speed INT,
        growth INT,
        gold VARCHAR(100),
        resources VARCHAR(100),
        ability VARCHAR(100),
        ai_value INT,
        img_data TEXT,
        PRIMARY KEY (`id`)
    )";
 $conn->exec($sql);

} catch(PDOException $e) {
  echo $sql . "<br>" . $e->getMessage();
}

foreach($result as $key => $value){
    
    $sql = "INSERT INTO units(name, level, fraction, min_damage, max_damage, attack, defence, hp, speed, growth, gold, resources, ability, ai_value, img_data)
                        VALUES(:name, :level, :fraction, :min_damage, :max_damage, :attack, :defence , :hp, :speed , :growth, :gold, :resources, :ability, :ai_value, :img_data)";
    $sql = $conn->prepare($sql);
    $sql->execute(array(
        ":name" => $value['name'],
        ":level" => $value['level'],
        ":fraction" => $value['fraction'],
        ":min_damage" => $value['min_damage'],
        ":max_damage" => $value['max_damage'],
        ":attack" => $value['attack'],
        ":defence" => $value['defence'],
        ":hp" => $value['hp'],
        ":speed" => $value['speed'],
        ":growth" => $value['growth'],
        ":gold" => $value['gold'],
        ":resources" => $value['resources'],
        ":ability" => $value['ability'],
        ":ai_value" => $value['ai_value'],
        ":img_data" => $value['img_data'],
    ));
}
