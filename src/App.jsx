/* eslint-disable */
import React, { useState, useMemo, useRef } from "react";

import * as XLSX from "xlsx";

const CATALOG = [{"id":"065-500","name":"Bolígrafo 6 colores en 1 Avengers sp319","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-504","name":"Bolígrafo 6 colores en 1 Batman ARTlj558","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-507","name":"Bolígrafo 6 colores en 1 Boca Juniors bo438","costPrice":2053.24,"salePrice":3079.86,"category":"Bolígrafos","stock":0},{"id":"065-509","name":"Bolígrafo 6 colores en 1 Cresko ARTck839","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":5},{"id":"065-495","name":"Bolígrafo 6 colores en 1 Harry potter ARThp419","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":5},{"id":"065-494","name":"Bolígrafo 6 colores en 1 La Liga de la justicia ARTlj546","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-502","name":"Bolígrafo 6 colores en 1 La liga de la justicia kids","costPrice":2053.24,"salePrice":3079.86,"category":"Bolígrafos","stock":0},{"id":"065-492","name":"Bolígrafo 6 colores en 1 Minecraft ARTmi412","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-506","name":"Bolígrafo 6 colores en 1 River plate ARTri505","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-505","name":"Bolígrafo 6 colores en 1 Simones ARTsi64/si757","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-497","name":"Bolígrafo 6 colores en 1 Sirenita ART ep776","costPrice":2053.24,"salePrice":3079.86,"category":"Bolígrafos","stock":0},{"id":"065-491","name":"Bolígrafo 6 colores en 1 Sonic ARTso461","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-459","name":"Bolígrafo 6 colores en 1 Spiderman ART512","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"065-499","name":"Bolígrafo 6 colores en 1 Spiderman ha483","costPrice":2053.24,"salePrice":3079.86,"category":"Bolígrafos","stock":5},{"id":"065-501","name":"Bolígrafo 6 colores en 1 Spidey ARTsp122","costPrice":2395.45,"salePrice":3593.17,"category":"Bolígrafos","stock":5},{"id":"065-508","name":"Bolígrafo 6 colores en 1 Tortugas ninja ARTtn121","costPrice":1960.74,"salePrice":2941.11,"category":"Bolígrafos","stock":0},{"id":"141-011","name":"Bolígrafo Bic 1 mm azul","costPrice":534.44,"salePrice":801.66,"category":"Bolígrafos","stock":0},{"id":"141-012","name":"Bolígrafo Bic 1 mm negro","costPrice":534.44,"salePrice":801.66,"category":"Bolígrafos","stock":0},{"id":"141-013","name":"Bolígrafo Bic 1 mm rojo","costPrice":534.44,"salePrice":801.66,"category":"Bolígrafos","stock":5},{"id":"141-014","name":"Bolígrafo Bic 1 mm verde","costPrice":534.44,"salePrice":801.66,"category":"Bolígrafos","stock":0},{"id":"141-162","name":"Bolígrafo Bic 4 colores blister (3 + 1 minas hb)","costPrice":4076.08,"salePrice":6114.12,"category":"Bolígrafos","stock":0},{"id":"141-165","name":"Bolígrafo Bic 4 colores blister fine","costPrice":3223.16,"salePrice":4834.74,"category":"Bolígrafos","stock":5},{"id":"141-160","name":"Bolígrafo Bic 4 colores blister firework","costPrice":3223.16,"salePrice":4834.74,"category":"Bolígrafos","stock":25},{"id":"141-164","name":"Bolígrafo Bic 4 colores blister flora","costPrice":3223.16,"salePrice":2417.37,"category":"Bolígrafos","stock":10},{"id":"141-163","name":"Bolígrafo Bic 4 colores blister fun","costPrice":3223.16,"salePrice":4834.74,"category":"Bolígrafos","stock":0},{"id":"141-161","name":"Bolígrafo Bic 4 colores blister highlighter fluo","costPrice":3223.16,"salePrice":4834.74,"category":"Bolígrafos","stock":0},{"id":"141-166","name":"Bolígrafo Bic 4 colores blister Pastel","costPrice":3223.16,"salePrice":4834.74,"category":"Bolígrafos","stock":20},{"id":"141-167","name":"Bolígrafo Bic 4 colores blister solar","costPrice":3223.16,"salePrice":4834.74,"category":"Bolígrafos","stock":15},{"id":"141-146","name":"Bolígrafo Bic cristal 0,7 blister x 4","costPrice":3590.29,"salePrice":4039.08,"category":"Bolígrafos","stock":10},{"id":"141-144","name":"Bolígrafo Bic cristal 0,7 blister x 8","costPrice":5498.98,"salePrice":8248.47,"category":"Bolígrafos","stock":0},{"id":"141-145","name":"Bolígrafo Bic cristal 1,6 blister x 10","costPrice":7551.74,"salePrice":11327.61,"category":"Bolígrafos","stock":5},{"id":"141-021","name":"Bolígrafo Bic cristal azul","costPrice":534.44,"salePrice":801.66,"category":"Bolígrafos","stock":15},{"id":"141-032","name":"Bolígrafo Bic cristal azul blister x 3","costPrice":2169.31,"salePrice":2440.47,"category":"Bolígrafos","stock":0},{"id":"141-027","name":"Bolígrafo Bic cristal blister x4 colores surtidos","costPrice":2547.97,"salePrice":3821.95,"category":"Bolígrafos","stock":0},{"id":"141-141","name":"Bolígrafo bic cristal fina 08 azul","costPrice":614.78,"salePrice":922.17,"category":"Bolígrafos","stock":20},{"id":"141-140","name":"Bolígrafo Bic cristal fina 08 blister x 3","costPrice":2520.99,"salePrice":3781.48,"category":"Bolígrafos","stock":0},{"id":"141-142","name":"Bolígrafo Bic cristal fina 08 negro","costPrice":614.78,"salePrice":922.17,"category":"Bolígrafos","stock":15},{"id":"141-143","name":"Bolígrafo Bic cristal fina 08 rojo","costPrice":614.78,"salePrice":922.17,"category":"Bolígrafos","stock":15},{"id":"141-030","name":"Bolígrafo Bic cristal intenso 1,6 azul","costPrice":755.54,"salePrice":1133.31,"category":"Bolígrafos","stock":10},{"id":"141-031","name":"Bolígrafo Bic cristal intenso 1,6 negro","costPrice":755.54,"salePrice":1133.31,"category":"Bolígrafos","stock":0},{"id":"096-101","name":"Accesorio para cuaderno inteligente dijes colgantes","costPrice":3850.7,"salePrice":5776.05,"category":"Cuadernos","stock":25},{"id":"096-102","name":"Accesorio para cuaderno inteligente disco grande","costPrice":4077.22,"salePrice":6115.83,"category":"Cuadernos","stock":30},{"id":"096-103","name":"Accesorio para cuaderno inteligente disco grande deluxe","costPrice":11368.07,"salePrice":17052.1,"category":"Cuadernos","stock":0},{"id":"232-927","name":"Caratula para cuaderno A5 Mooving loop Harry potter ART 1724222","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":20},{"id":"232-922","name":"Caratula para cuaderno A5 Mooving loop Hello kitty ART1724165","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":0},{"id":"232-928","name":"Caratula para cuaderno A5 Mooving loop Mickey mouse ART 1724121","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":30},{"id":"232-929","name":"Caratula para cuaderno A5 Mooving loop Minnie mouse ART 1724131","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":10},{"id":"232-923","name":"Caratula para cuaderno A5 Mooving loop Pusheen ART 1724207","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":15},{"id":"232-924","name":"Caratula para cuaderno A5 Mooving loop Snoopy ART 1724134","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":5},{"id":"232-919","name":"Caratula para cuaderno A5 Mooving loop Stitch ART 1724244","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":0},{"id":"232-659","name":"Caratula para cuaderno A5 Mooving loop winnie the pooh ART 1724122","costPrice":8439.15,"salePrice":12658.72,"category":"Cuadernos","stock":0},{"id":"232-865","name":"Caratula para cuaderno carta Harry potter Mooving loop ART 1714222","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":5},{"id":"232-866","name":"Caratula para cuaderno carta Mickey mouse Mooving loop ART 1714121","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":10},{"id":"232-867","name":"Caratula para cuaderno carta Minnie mouse Mooving loop ART 1714131","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":0},{"id":"232-660","name":"Caratula para cuaderno carta Mooving loop Hello Kitty ART 1714165","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":5},{"id":"232-661","name":"Caratula para cuaderno carta Mooving loop winnie the pooh ART 1714122","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":0},{"id":"232-868","name":"Caratula para cuaderno carta organizador Mooving loop ART 1713132","costPrice":10491.91,"salePrice":15737.86,"category":"Cuadernos","stock":20},{"id":"232-920","name":"Caratula para cuaderno carta Pusheen Mooving loop ART1714207","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":10},{"id":"232-921","name":"Caratula para cuaderno carta Snoopy Mooving loop ART1714134","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":25},{"id":"232-869","name":"Caratula para cuaderno carta Stitch Mooving loop ART 1714244","costPrice":10263.83,"salePrice":15395.74,"category":"Cuadernos","stock":15},{"id":"233-950","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Afa","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":0},{"id":"233-963","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral boho","costPrice":5245.96,"salePrice":7868.94,"category":"Cuadernos","stock":15},{"id":"233-964","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral garden","costPrice":5245.96,"salePrice":7868.94,"category":"Cuadernos","stock":15},{"id":"233-951","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Harry potter","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":5},{"id":"233-952","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Hello Kitty","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":10},{"id":"233-953","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Kuromi","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":0},{"id":"233-954","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Lotso","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":0},{"id":"233-955","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Mickey and Friends","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":30},{"id":"233-956","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Minnie","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":5},{"id":"233-957","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral My melody","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":0},{"id":"233-958","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Pastel block","costPrice":5245.96,"salePrice":7868.94,"category":"Cuadernos","stock":25},{"id":"233-959","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral pink","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":20},{"id":"233-961","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Snoopy","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":10},{"id":"233-962","name":"Cuaderno 12 x 14,5 Mooving pocket premium tapa dura 96 hojas rayado espiral Stitch","costPrice":5702.13,"salePrice":8553.19,"category":"Cuadernos","stock":30},{"id":"122-326","name":"Cuaderno 16 x 21 América colección 80 hojas rayado espiral","costPrice":2852.74,"salePrice":4279.11,"category":"Cuadernos","stock":5},{"id":"122-262","name":"Cuaderno 16 x 21 Arte bloom tapa dura 80 hojas rayado espiral","costPrice":6575.33,"salePrice":9862.99,"category":"Cuadernos","stock":15},{"id":"122-703","name":"Cuaderno 16 x 21 Arte minimalista tapa semirígida 80 hojas rayado espiral","costPrice":6575.49,"salePrice":9863.24,"category":"Cuadernos","stock":0},{"id":"122-235","name":"Cuaderno 16 x 21 Arte music tapa dura 80 hojas rayado espiral","costPrice":8075.16,"salePrice":12112.74,"category":"Cuadernos","stock":5},{"id":"122-700","name":"Cuaderno 16 x 21 Arte nordico tapa dura 80 hojas rayado espiral","costPrice":8936.47,"salePrice":13404.7,"category":"Cuadernos","stock":0},{"id":"125-098","name":"Cuaderno 16 x 21 Avon 46 hojas rayado espiral","costPrice":1607.5,"salePrice":2411.25,"category":"Cuadernos","stock":15},{"id":"232-099","name":"Banderitas plásticas Mooving angostas resaltadoras 14 x 5 mm ART 2112100208","costPrice":3210.3,"salePrice":4815.45,"category":"Resaltadores","stock":20},{"id":"069-208","name":"Exhibidor Trabi resaltadores hight fluo x 128","costPrice":66085.12,"salePrice":99127.68,"category":"Resaltadores","stock":10},{"id":"069-207","name":"Exhibidor Trabi resaltadores Pastel x 128","costPrice":68532.32,"salePrice":102798.48,"category":"Resaltadores","stock":0},{"id":"141-716","name":"Resaltador Bic intensity doble punta x6","costPrice":9174.33,"salePrice":13761.49,"category":"Resaltadores","stock":5},{"id":"229-300","name":"Resaltador Carioca Pastel caja x 4","costPrice":7221.36,"salePrice":10832.04,"category":"Resaltadores","stock":15},{"id":"034-708","name":"Resaltador doble punta Filgo multilayer x6 Surtido","costPrice":4813.0,"salePrice":7219.5,"category":"Resaltadores","stock":5},{"id":"034-710","name":"Resaltador doble punta Filgo multilighter x10 fluo","costPrice":8668.8,"salePrice":13003.2,"category":"Resaltadores","stock":25},{"id":"034-711","name":"Resaltador doble punta Filgo multilighter x10 Pastel","costPrice":8668.8,"salePrice":13003.2,"category":"Resaltadores","stock":20},{"id":"040-530","name":"Resaltador Ezco fino amarillo","costPrice":301.71,"salePrice":407.31,"category":"Resaltadores","stock":25},{"id":"040-531","name":"Resaltador Ezco fino celeste","costPrice":301.71,"salePrice":407.31,"category":"Resaltadores","stock":0},{"id":"040-532","name":"Resaltador Ezco fino naranja","costPrice":301.71,"salePrice":407.31,"category":"Resaltadores","stock":10},{"id":"040-534","name":"Resaltador Ezco fino rosa","costPrice":301.71,"salePrice":407.31,"category":"Resaltadores","stock":0},{"id":"040-533","name":"Resaltador Ezco fino verde","costPrice":301.71,"salePrice":407.31,"category":"Resaltadores","stock":5},{"id":"134-610","name":"Resaltador Faber-castell t48 amarillo","costPrice":599.02,"salePrice":898.53,"category":"Resaltadores","stock":30},{"id":"034-040","name":"Resaltador Filgo lighter fine amarillo","costPrice":355.63,"salePrice":533.44,"category":"Resaltadores","stock":30},{"id":"034-044","name":"Resaltador Filgo lighter fine celeste","costPrice":355.63,"salePrice":533.44,"category":"Resaltadores","stock":10},{"id":"034-078","name":"Resaltador Filgo lighter fine estuche x18 Pastel","costPrice":8425.31,"salePrice":12637.97,"category":"Resaltadores","stock":20},{"id":"034-050","name":"Resaltador Filgo lighter fine exhibidor x36","costPrice":12804.46,"salePrice":19206.69,"category":"Resaltadores","stock":20},{"id":"034-051","name":"Resaltador Filgo lighter fine exhibidor x72 fluo","costPrice":31327.4,"salePrice":46991.1,"category":"Resaltadores","stock":15},{"id":"034-052","name":"Resaltador Filgo lighter fine exhibidor x72 Pastel","costPrice":34175.34,"salePrice":51263.01,"category":"Resaltadores","stock":5},{"id":"034-049","name":"Resaltador Filgo lighter fine fluo x10","costPrice":4138.19,"salePrice":6207.28,"category":"Resaltadores","stock":0},{"id":"034-047","name":"Resaltador Filgo lighter fine garden x6","costPrice":1380.09,"salePrice":2070.13,"category":"Resaltadores","stock":30},{"id":"034-042","name":"Resaltador Filgo lighter fine naranja","costPrice":355.63,"salePrice":533.44,"category":"Resaltadores","stock":25},{"id":"034-060","name":"Resaltador Filgo lighter fine Pastel amarillo","costPrice":355.74,"salePrice":533.61,"category":"Resaltadores","stock":0},{"id":"034-061","name":"Resaltador Filgo lighter fine Pastel celeste","costPrice":355.74,"salePrice":533.61,"category":"Resaltadores","stock":0},{"id":"034-075","name":"Resaltador Filgo lighter fine Pastel exhibidor x36","costPrice":12804.46,"salePrice":19206.69,"category":"Resaltadores","stock":0},{"id":"034-063","name":"Resaltador Filgo lighter fine Pastel naranja","costPrice":355.74,"salePrice":533.61,"category":"Resaltadores","stock":0},{"id":"034-064","name":"Resaltador Filgo lighter fine Pastel rosa","costPrice":355.74,"salePrice":533.61,"category":"Resaltadores","stock":0},{"id":"034-062","name":"Resaltador Filgo lighter fine Pastel verde","costPrice":355.74,"salePrice":533.61,"category":"Resaltadores","stock":20},{"id":"034-065","name":"Resaltador Filgo lighter fine Pastel violeta","costPrice":355.74,"salePrice":533.61,"category":"Resaltadores","stock":0},{"id":"034-048","name":"Resaltador Filgo lighter fine Pastel x10","costPrice":4138.19,"salePrice":6207.28,"category":"Resaltadores","stock":20},{"id":"034-070","name":"Resaltador Filgo lighter fine Pastel x4","costPrice":1653.74,"salePrice":2480.61,"category":"Resaltadores","stock":20},{"id":"034-043","name":"Resaltador Filgo lighter fine rosa","costPrice":355.63,"salePrice":533.44,"category":"Resaltadores","stock":25},{"id":"034-041","name":"Resaltador Filgo lighter fine verde","costPrice":355.63,"salePrice":533.44,"category":"Resaltadores","stock":30},{"id":"034-045","name":"Resaltador Filgo lighter fine violeta","costPrice":355.63,"salePrice":533.44,"category":"Resaltadores","stock":10},{"id":"034-001","name":"Resaltador Filgo textmarker amarillo","costPrice":606.13,"salePrice":909.19,"category":"Resaltadores","stock":30},{"id":"034-005","name":"Resaltador Filgo textmarker celeste","costPrice":606.13,"salePrice":909.19,"category":"Resaltadores","stock":0},{"id":"034-008","name":"Resaltador Filgo textmarker exhibidor x60 fluo","costPrice":41342.35,"salePrice":62013.52,"category":"Resaltadores","stock":0},{"id":"034-024","name":"Resaltador Filgo textmarker fluo x10","costPrice":7045.83,"salePrice":10568.74,"category":"Resaltadores","stock":30},{"id":"034-059","name":"Resaltador Filgo textmarker mute blister x 5","costPrice":3480.21,"salePrice":5220.32,"category":"Resaltadores","stock":10},{"id":"146-378","name":"Block Artmate p/ marcadores A4 100g pegado 24 hojas","costPrice":4654.1,"salePrice":6981.15,"category":"Marcadores","stock":15},{"id":"085-010","name":"Borrador para pizarra blanca con repuesto y portamarcador","costPrice":1565.81,"salePrice":2348.72,"category":"Marcadores","stock":0},{"id":"166-041","name":"Borrador para pizarra Sifap con iman y porta marcadores","costPrice":2418.83,"salePrice":3628.24,"category":"Marcadores","stock":10},{"id":"145-128","name":"Compas pizarra natural para marcador","costPrice":20073.8,"salePrice":30110.7,"category":"Marcadores","stock":20},{"id":"034-860","name":"Exhibidor marcador Filgo para pizarra x144 unidades 058","costPrice":148879.13,"salePrice":223318.7,"category":"Marcadores","stock":0},{"id":"034-830","name":"Exhibidor marcador Filgo permanente x144 050 051","costPrice":126997.69,"salePrice":190496.54,"category":"Marcadores","stock":25},{"id":"069-470","name":"Exhibidor Trabi marcador 410/411 permanente x 128 unidades","costPrice":87123.93,"salePrice":130685.89,"category":"Marcadores","stock":0},{"id":"069-472","name":"Exhibidor Trabi marcador 420 al agua x 128 unidades","costPrice":70215.44,"salePrice":105323.16,"category":"Marcadores","stock":10},{"id":"141-770","name":"Kit marcadores tonos Pastel Bic x 12 unidades","costPrice":15635.37,"salePrice":23453.06,"category":"Marcadores","stock":30},{"id":"141-781","name":"Kit marcadores tonos Pastel Bic x 6 unidades","costPrice":6324.08,"salePrice":9486.12,"category":"Marcadores","stock":0},{"id":"151-947","name":"Lapices de colores infinity x 15 + marcadores Ocean x 12","costPrice":9132.37,"salePrice":13698.56,"category":"Marcadores","stock":30},{"id":"082-816","name":"Lápiz doms mega triangular 12 colores largos c/ sacapuntas+2 marcadores","costPrice":7812.38,"salePrice":11718.57,"category":"Marcadores","stock":0},{"id":"064-437","name":"Libro para colorear cajitas magicas c/marcadores colores magicos","costPrice":19199.4,"salePrice":28799.1,"category":"Marcadores","stock":10},{"id":"064-438","name":"Libro para colorear cajitas magicas c/marcadores deseos estelares","costPrice":19199.4,"salePrice":28799.1,"category":"Marcadores","stock":30},{"id":"141-730","name":"Marcador Bic evolution plumones x 12 colores","costPrice":3107.2,"salePrice":4660.8,"category":"Marcadores","stock":5},{"id":"141-700","name":"Marcador Bic permanent marking grip su bl x 36","costPrice":48655.38,"salePrice":36491.53,"category":"Marcadores","stock":0},{"id":"141-704","name":"Marcador Bic permanente grip bl x12","costPrice":17285.37,"salePrice":12964.02,"category":"Marcadores","stock":15},{"id":"141-703","name":"Marcador Bic permanente grip bl x6 fashion","costPrice":9457.17,"salePrice":7092.87,"category":"Marcadores","stock":0},{"id":"141-702","name":"Marcador Bic permanente grip bl x6 pasteles","costPrice":9241.39,"salePrice":6931.05,"category":"Marcadores","stock":30},{"id":"141-710","name":"Marcador Bic permanente grip metalizado x4","costPrice":6096.67,"salePrice":9145.01,"category":"Marcadores","stock":30},{"id":"082-820","name":"Marcador doms acuarelable 12 colores","costPrice":1665.3,"salePrice":2497.95,"category":"Marcadores","stock":0},{"id":"082-821","name":"Marcador doms acuarelable 24 colores","costPrice":3341.78,"salePrice":5012.67,"category":"Marcadores","stock":15},{"id":"082-825","name":"Marcador doms jumbo 12 colores (cuerpo decorado)","costPrice":4459.43,"salePrice":6689.15,"category":"Marcadores","stock":25},{"id":"082-828","name":"Marcador doms punta pincel 13 colores + 1 difuminador","costPrice":8930.03,"salePrice":13395.05,"category":"Marcadores","stock":0},{"id":"102-041","name":"Marcador Edding 100 permanente punta redonda azul","costPrice":1697.17,"salePrice":2545.76,"category":"Marcadores","stock":0},{"id":"102-040","name":"Marcador Edding 100 permanente punta redonda negro","costPrice":1697.17,"salePrice":2545.76,"category":"Marcadores","stock":15},{"id":"102-042","name":"Marcador Edding 100 permanente punta redonda rojo","costPrice":1697.17,"salePrice":2545.76,"category":"Marcadores","stock":10},{"id":"102-044","name":"Marcador Edding 100 permanente punta redonda verde","costPrice":1697.17,"salePrice":2545.76,"category":"Marcadores","stock":5},{"id":"102-035","name":"Marcador Edding 1200 punta fina amarillo","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":0},{"id":"102-001","name":"Marcador Edding 1200 punta fina azul","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":5},{"id":"102-031","name":"Marcador Edding 1200 punta fina celeste","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":0},{"id":"102-034","name":"Marcador Edding 1200 punta fina naranja","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":0},{"id":"102-002","name":"Marcador Edding 1200 punta fina negro","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":25},{"id":"102-003","name":"Marcador Edding 1200 punta fina rojo","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":0},{"id":"102-030","name":"Marcador Edding 1200 punta fina rosa","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":30},{"id":"102-004","name":"Marcador Edding 1200 punta fina verde","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":0},{"id":"102-036","name":"Marcador Edding 1200 punta fina verde claro","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":0},{"id":"102-032","name":"Marcador Edding 1200 punta fina violeta","costPrice":1772.99,"salePrice":2659.49,"category":"Marcadores","stock":25},{"id":"102-070","name":"Marcador Edding 130 permanente punta chanfle negro","costPrice":1697.17,"salePrice":2545.76,"category":"Marcadores","stock":30},{"id":"102-451","name":"Marcador Edding 160 para pizarra azul","costPrice":1789.71,"salePrice":2684.57,"category":"Marcadores","stock":0},{"id":"030-930","name":"Adhesivo en barra Sifap lápiz 6 gramos","costPrice":1119.25,"salePrice":1678.88,"category":"Lápices","stock":10},{"id":"060-839","name":"Agua wow! libro c/lapiz magico para pintar c/agua animales","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":30},{"id":"060-835","name":"Agua wow! libro c/lapiz magico para pintar c/agua Bluey","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":20},{"id":"060-830","name":"Agua wow! libro c/lapiz magico para pintar c/agua Frozen","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":5},{"id":"060-829","name":"Agua wow! libro c/lapiz magico para pintar c/agua Hello kitty","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":30},{"id":"060-828","name":"Agua wow! libro c/lapiz magico para pintar c/agua Jurassic World","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":5},{"id":"060-836","name":"Agua wow! libro c/lapiz magico para pintar c/agua ositos cariñositos","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":10},{"id":"060-834","name":"Agua wow! libro c/lapiz magico para pintar c/agua Sonic","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":20},{"id":"060-838","name":"Agua wow! libro c/lapiz magico para pintar c/agua Spidey","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":15},{"id":"060-840","name":"Agua wow! libro c/lapiz magico para pintar c/agua Toy Story","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":25},{"id":"060-842","name":"Agua wow! libro c/lapiz magico para pintar c/agua transportes","costPrice":6183.5,"salePrice":9275.25,"category":"Lápices","stock":30},{"id":"232-060","name":"Banderitas plásticas Mooving lápiz neon x 125 ART2112100203","costPrice":1912.49,"salePrice":2868.74,"category":"Lápices","stock":25},{"id":"017-295","name":"Cartulina magica 14 x 20 cm 5 hojas + lápiz magico mix","costPrice":2394.94,"salePrice":3592.41,"category":"Lápices","stock":0},{"id":"017-296","name":"Cartulina magica 20 x 28 cm 5 hojas + lápiz magico mix","costPrice":3625.75,"salePrice":5438.62,"category":"Lápices","stock":5},{"id":"151-258","name":"Compas Maped study universal p/lapiz c/abrazadera sy 120","costPrice":3616.82,"salePrice":5425.23,"category":"Lápices","stock":5},{"id":"143-005","name":"Compas Pizzini linear ART122 metal lápiz","costPrice":3428.08,"salePrice":5142.12,"category":"Lápices","stock":0},{"id":"034-281","name":"Compas plástico Filgo giro con lápiz","costPrice":1116.5,"salePrice":1674.75,"category":"Lápices","stock":15},{"id":"034-780","name":"Goma de borrar Filgo lápiz tek 4080 x 36","costPrice":10236.91,"salePrice":15355.36,"category":"Lápices","stock":0},{"id":"040-996","name":"Goma Ezco caucho plástico lápiz ral40 blanca caja x40","costPrice":6718.3,"salePrice":9069.7,"category":"Lápices","stock":30},{"id":"040-994","name":"Goma Ezco caucho plástico tinta/lapiz br40 azul/roja caja x 40","costPrice":8117.16,"salePrice":10958.16,"category":"Lápices","stock":5},{"id":"040-995","name":"Goma Ezco caucho plástico tinta/lapiz dg40 gris caja x40","costPrice":6684.22,"salePrice":9023.7,"category":"Lápices","stock":5},{"id":"040-714","name":"Goma Ezco técnica plástica caja x 30 lápiz","costPrice":5923.51,"salePrice":7996.74,"category":"Lápices","stock":0},{"id":"040-715","name":"Goma Ezco técnica plástica caja x 30 lapiz/tinta","costPrice":6842.94,"salePrice":9237.97,"category":"Lápices","stock":0},{"id":"134-154","name":"Goma Faber-castell 7082 caja x 30 lapiz/tinta","costPrice":16849.6,"salePrice":25274.4,"category":"Lápices","stock":0},{"id":"134-157","name":"Goma Faber-castell 7086 caja x 30 lápiz","costPrice":14689.17,"salePrice":22033.76,"category":"Lápices","stock":5},{"id":"041-320","name":"Goma Staedtler 526 x30 lápiz","costPrice":17891.35,"salePrice":26837.02,"category":"Lápices","stock":0},{"id":"115-265","name":"Gomas lápiz staedler c/escob rasor","costPrice":2457.66,"salePrice":3686.49,"category":"Lápices","stock":0},{"id":"134-275","name":"Lapices de colores Faber-castell ecolapiz acuarel x 12 largo caja cartón","costPrice":10435.32,"salePrice":15652.98,"category":"Lápices","stock":15},{"id":"134-276","name":"Lapices de colores Faber-castell ecolapiz acuarel x 24 largo caja cartón","costPrice":19740.85,"salePrice":29611.27,"category":"Lápices","stock":0},{"id":"134-277","name":"Lapices de colores Faber-castell ecolapiz acuarel x 36 largo caja cartón","costPrice":29254.38,"salePrice":43881.57,"category":"Lápices","stock":30},{"id":"040-350","name":"Carpeta 2 x 40 mm A4 transparente Ezco neon Surtido","costPrice":2660.39,"salePrice":3591.52,"category":"Carpetas","stock":5},{"id":"040-351","name":"Carpeta 2 x 40 mm oficio transparente Ezco neon Surtido","costPrice":2733.76,"salePrice":3690.57,"category":"Carpetas","stock":10},{"id":"039-587","name":"Carpeta 2-a/red,40 mm A4 forrada azul Lama","costPrice":2758.81,"salePrice":4138.22,"category":"Carpetas","stock":25},{"id":"039-583","name":"Carpeta 2-a/red,40 mm A4 forrada naranja Lama","costPrice":2758.81,"salePrice":4138.22,"category":"Carpetas","stock":5},{"id":"039-585","name":"Carpeta 2-a/red,40 mm A4 forrada purpura Lama","costPrice":2758.81,"salePrice":4138.22,"category":"Carpetas","stock":30},{"id":"039-581","name":"Carpeta 2-a/red,40 mm A4 forrada rojo Lama","costPrice":2758.81,"salePrice":4138.22,"category":"Carpetas","stock":0},{"id":"039-584","name":"Carpeta 2-a/red,40 mm A4 forrada v, manzana Lama","costPrice":2758.81,"salePrice":4138.22,"category":"Carpetas","stock":25},{"id":"039-582","name":"Carpeta 2-a/red,40 mm A4 forrada verde Lama","costPrice":2758.81,"salePrice":4138.22,"category":"Carpetas","stock":5},{"id":"017-625","name":"Carpeta 2-a/red,40 mm cartoné Reysa Argentina","costPrice":5445.0,"salePrice":8167.5,"category":"Carpetas","stock":25},{"id":"017-628","name":"Carpeta 2-a/red,40 mm cartoné Reysa Pastel","costPrice":5445.0,"salePrice":8167.5,"category":"Carpetas","stock":20},{"id":"017-626","name":"Carpeta 2-a/red,40 mm cartoné Reysa street","costPrice":5445.0,"salePrice":8167.5,"category":"Carpetas","stock":5},{"id":"017-627","name":"Carpeta 2-a/red,40 mm cartoné Reysa Trendy","costPrice":5445.0,"salePrice":8167.5,"category":"Carpetas","stock":0},{"id":"060-578","name":"Carpeta 3 anillos redondos 40 mm cartoné Big life Boca","costPrice":7207.49,"salePrice":10811.24,"category":"Carpetas","stock":0},{"id":"060-571","name":"Carpeta 3 anillos redondos 40 mm cartoné Big life Just cool","costPrice":5295.78,"salePrice":7943.67,"category":"Carpetas","stock":20},{"id":"060-567","name":"Carpeta 3 anillos redondos 40 mm cartoné Big life Sonic","costPrice":7207.49,"salePrice":10811.24,"category":"Carpetas","stock":15},{"id":"122-620","name":"Carpeta 3 anillos redondos 40 mm cartoné rivad+ repuesto 96 hojas ry+rep48 hojas cuadriculado","costPrice":19088.54,"salePrice":28632.81,"category":"Carpetas","stock":20},{"id":"039-561","name":"Carpeta 3 anillos redondos 40 mm forrada azul Lama","costPrice":2338.12,"salePrice":3507.18,"category":"Carpetas","stock":20},{"id":"039-564","name":"Carpeta 3 anillos redondos 40 mm forrada naranja Lama","costPrice":2338.12,"salePrice":3507.18,"category":"Carpetas","stock":25},{"id":"039-567","name":"Carpeta 3 anillos redondos 40 mm forrada negro Lama","costPrice":2338.12,"salePrice":3507.18,"category":"Carpetas","stock":0},{"id":"039-566","name":"Carpeta 3 anillos redondos 40 mm forrada purpura Lama","costPrice":2338.12,"salePrice":3507.18,"category":"Carpetas","stock":0},{"id":"039-562","name":"Carpeta 3 anillos redondos 40 mm forrada rojo Lama","costPrice":2338.12,"salePrice":3507.18,"category":"Carpetas","stock":0},{"id":"039-565","name":"Carpeta 3 anillos redondos 40 mm forrada v manzana Lama","costPrice":2338.12,"salePrice":3507.18,"category":"Carpetas","stock":20},{"id":"136-112","name":"Carpeta 3 anillos redondos 40 mm forrada verde","costPrice":2420.24,"salePrice":3630.36,"category":"Carpetas","stock":15},{"id":"039-563","name":"Carpeta 3 anillos redondos 40 mm forrada verde Lama","costPrice":2338.12,"salePrice":3507.18,"category":"Carpetas","stock":0},{"id":"231-080","name":"Carpeta 3 anillos redondos 40 mm Mooving cartoné Afa","costPrice":9123.4,"salePrice":13685.1,"category":"Carpetas","stock":5},{"id":"231-086","name":"Carpeta 3 anillos redondos 40 mm Mooving cartoné balloon","costPrice":6728.51,"salePrice":10092.76,"category":"Carpetas","stock":5},{"id":"231-105","name":"Carpeta 3 anillos redondos 40 mm Mooving cartoné boho","costPrice":7982.98,"salePrice":11974.47,"category":"Carpetas","stock":5},{"id":"231-102","name":"Carpeta 3 anillos redondos 40 mm Mooving cartoné Como Quieres","costPrice":9123.4,"salePrice":13685.1,"category":"Carpetas","stock":30},{"id":"231-115","name":"Carpeta 3 anillos redondos 40 mm Mooving cartoné cool","costPrice":7412.76,"salePrice":11119.14,"category":"Carpetas","stock":25},{"id":"231-123","name":"Carpeta 3 anillos redondos 40 mm Mooving cartoné dc","costPrice":9123.4,"salePrice":13685.1,"category":"Carpetas","stock":0},{"id":"192-419","name":"Ad-astra ficha cte cte 2065f paqx100","costPrice":5436.23,"salePrice":8154.34,"category":"Blocks","stock":20},{"id":"192-153","name":"Ad-astra ficha -legajo personal 2055f-","costPrice":21173.84,"salePrice":31760.76,"category":"Blocks","stock":0},{"id":"017-145","name":"Anotador con solapa Reysa 15 x 20 cm 50 hojas rayadas mix","costPrice":4729.24,"salePrice":7093.86,"category":"Blocks","stock":10},{"id":"111-001","name":"Anotador Congreso 40 hs rayado","costPrice":2045.29,"salePrice":3067.93,"category":"Blocks","stock":25},{"id":"232-983","name":"Anotador Mooving Lotso 10 cm x 20 cm ART 2541257001","costPrice":7687.6,"salePrice":11531.4,"category":"Blocks","stock":5},{"id":"017-144","name":"Anotadores block con imán Reysa 9 x 15 50 hs listas","costPrice":1198.07,"salePrice":1797.11,"category":"Blocks","stock":0},{"id":"111-044","name":"Block A4 perforado Mis apuntes 80 hojas cuadriculado","costPrice":3516.49,"salePrice":5274.73,"category":"Blocks","stock":25},{"id":"111-043","name":"Block A4 perforado Mis apuntes 80 hs rayado","costPrice":3516.49,"salePrice":5274.73,"category":"Blocks","stock":30},{"id":"125-370","name":"Block anotador Avon A5 80 hojas rayado espiral","costPrice":3166.32,"salePrice":4749.48,"category":"Blocks","stock":0},{"id":"125-372","name":"Block anotador Avon A5 80 hs liso espiral","costPrice":3166.32,"salePrice":4749.48,"category":"Blocks","stock":0},{"id":"125-364","name":"Block anotador Éxito 16 x 21 48 hojas rayado","costPrice":2031.69,"salePrice":3047.53,"category":"Blocks","stock":30},{"id":"146-380","name":"Block Artmate de hojas negras A4 140 gramos espiral 30 hojas","costPrice":6205.6,"salePrice":9308.4,"category":"Blocks","stock":0},{"id":"146-376","name":"Block Artmate p/ boceto A4 100g pegado 40 hojas","costPrice":5049.68,"salePrice":7574.52,"category":"Blocks","stock":0},{"id":"146-377","name":"Block Artmate p/ boceto A4 160 gramos espiral 30 hojas","costPrice":5829.56,"salePrice":8744.34,"category":"Blocks","stock":5},{"id":"146-374","name":"Block Artmate p/ carbonilla A4 160 gramos espiral 32 hojas","costPrice":6748.54,"salePrice":10122.81,"category":"Blocks","stock":0},{"id":"146-379","name":"Block Artmate p/ oleo A4 200 gramos pegado A4","costPrice":6671.05,"salePrice":10006.58,"category":"Blocks","stock":20},{"id":"146-375","name":"Block Artmate p/ tiza Pastel A4 180 gramos espiral 24 hojas","costPrice":9618.71,"salePrice":14428.06,"category":"Blocks","stock":25},{"id":"125-400","name":"Block Avon A4 perforado cuadriculado x 80 hs","costPrice":3601.35,"salePrice":5402.02,"category":"Blocks","stock":25},{"id":"125-399","name":"Block Avon A4 perforado rayado x 80 hs","costPrice":3601.35,"salePrice":5402.02,"category":"Blocks","stock":5},{"id":"125-426","name":"Block Avon borrador liso A4 x 80 hs","costPrice":3601.35,"salePrice":5402.02,"category":"Blocks","stock":20},{"id":"125-423","name":"Block Avon borrador liso A5 x 80 hs","costPrice":1917.62,"salePrice":2876.43,"category":"Blocks","stock":0},{"id":"125-420","name":"Block Avon borrador liso A6 x 80 hs","costPrice":1138.62,"salePrice":1707.93,"category":"Blocks","stock":0},{"id":"125-429","name":"Block Avon borrador liso oficio x 80 hs","costPrice":4151.73,"salePrice":6227.59,"category":"Blocks","stock":20},{"id":"125-406","name":"Block Avon oficio perforado cuadriculado x 80 hs","costPrice":4217.41,"salePrice":6326.11,"category":"Blocks","stock":0},{"id":"125-404","name":"Block Avon oficio perforado rayado x 80 hs","costPrice":4217.41,"salePrice":6326.11,"category":"Blocks","stock":20},{"id":"224-223","name":"Camioneta a friccion con ruedas de goma ART54570","costPrice":6102.71,"salePrice":9154.07,"category":"Gomas","stock":10},{"id":"266-052","name":"Canopla goma eva Cresko 1 piso chica Robots ARTck326","costPrice":19607.45,"salePrice":29411.18,"category":"Gomas","stock":25},{"id":"050-025","name":"Cartuchera Hermosura 1 cierre sofi uncornio goma","costPrice":5158.23,"salePrice":7737.34,"category":"Gomas","stock":10},{"id":"078-370","name":"Crayones para el agua Wero x6 con figuras goma eva Ocean ARTwe4191","costPrice":6374.98,"salePrice":9562.47,"category":"Gomas","stock":20},{"id":"117-032","name":"Didac Kreker 329 goma Kreker imantada formas y siluetas 47 pcs","costPrice":6938.24,"salePrice":10407.36,"category":"Gomas","stock":30},{"id":"117-220","name":"Didactico de goma eva reloj Kreker tiempo al tiempo 23 x 36cms ART1034","costPrice":7370.12,"salePrice":11055.18,"category":"Gomas","stock":25},{"id":"224-321","name":"Dinosaurio muñeco de goma 20cms ART 55187","costPrice":4478.06,"salePrice":6717.09,"category":"Gomas","stock":0},{"id":"117-100","name":"Encastre de goma eva foamy Kreker aprendemos en abacedario 34 x 25cms ART816","costPrice":7031.52,"salePrice":10547.28,"category":"Gomas","stock":5},{"id":"117-102","name":"Encastre de goma eva foamy Kreker aprendemos la hora 22 x 33cms ART819","costPrice":6742.74,"salePrice":10114.11,"category":"Gomas","stock":10},{"id":"117-105","name":"Encastre de goma eva foamy Kreker blocks 80 piezas ART1818","costPrice":8369.33,"salePrice":12553.99,"category":"Gomas","stock":5},{"id":"117-118","name":"Encastre de goma eva foamy Kreker mapa Argentina a lo grande 27 x 40cms ART818","costPrice":11608.45,"salePrice":17412.68,"category":"Gomas","stock":0},{"id":"117-115","name":"Encastre de goma eva foamy Kreker mapa Argentina animales de mi pais ART810","costPrice":8196.83,"salePrice":12295.24,"category":"Gomas","stock":30},{"id":"117-116","name":"Encastre de goma eva foamy Kreker mi cuerpo por dentro 22 x 33cms ART821","costPrice":6612.41,"salePrice":9918.61,"category":"Gomas","stock":0},{"id":"117-108","name":"Encastre de goma eva foamy Kreker mis emociones 23 x 36cms ART811","costPrice":8196.83,"salePrice":12295.24,"category":"Gomas","stock":15},{"id":"117-110","name":"Encastre de goma eva foamy puzzle Kreker mundo animal 22 x 36cms ART813","costPrice":6061.69,"salePrice":9092.53,"category":"Gomas","stock":0},{"id":"117-021","name":"Encastre de goma eva Kreker cuerpo humano mi cuerpo 22 x 33cms ART334","costPrice":4607.61,"salePrice":6911.41,"category":"Gomas","stock":0},{"id":"117-012","name":"Encastre de goma eva Kreker gracias agua 22 piezas ART227","costPrice":9698.2,"salePrice":14547.3,"category":"Gomas","stock":25},{"id":"117-009","name":"Encastre de goma eva Kreker mapa mi argetina 22 x 35cms ART320","costPrice":6502.52,"salePrice":9753.78,"category":"Gomas","stock":30},{"id":"117-005","name":"Encastre de goma eva Kreker mapa yo amo a mi pais 24 x 28cms ART314","costPrice":4357.16,"salePrice":6535.74,"category":"Gomas","stock":30},{"id":"117-010","name":"Encastre de goma eva Kreker mi abecedario 20 x 26cms ART343","costPrice":5136.6,"salePrice":7704.9,"category":"Gomas","stock":0},{"id":"155-408","name":"Asta reglamentaria - conforme a la norma","costPrice":130168.99,"salePrice":195253.49,"category":"Geometría","stock":0},{"id":"155-399","name":"Asta reglamentaria Jardin","costPrice":73383.09,"salePrice":110074.63,"category":"Geometría","stock":30},{"id":"155-340","name":"Bandera ceremonia reglamentaria Argentina","costPrice":65879.13,"salePrice":98818.7,"category":"Geometría","stock":0},{"id":"155-920","name":"Bandera ceremonia reglamentaria bonaerense (90*150)","costPrice":157640.95,"salePrice":236461.43,"category":"Geometría","stock":0},{"id":"232-074","name":"Banderitas plásticas Mooving neon 8 x 25 c/regla ART 2112100205","costPrice":2141.72,"salePrice":3212.58,"category":"Geometría","stock":0},{"id":"040-956","name":"Compas Ezco fiddo plástico","costPrice":720.99,"salePrice":973.34,"category":"Geometría","stock":0},{"id":"040-955","name":"Compas Ezco progression metálico","costPrice":1796.22,"salePrice":2424.9,"category":"Geometría","stock":5},{"id":"151-292","name":"Compas Maped kid' z","costPrice":1351.72,"salePrice":2027.58,"category":"Geometría","stock":20},{"id":"151-293","name":"Compas Maped kid' z c/abrazadera","costPrice":1496.31,"salePrice":2244.47,"category":"Geometría","stock":0},{"id":"151-571","name":"Compas Maped neon study abrazadera","costPrice":4132.46,"salePrice":6198.69,"category":"Geometría","stock":5},{"id":"151-570","name":"Compas Maped neon study mina","costPrice":3669.26,"salePrice":5503.89,"category":"Geometría","stock":0},{"id":"151-260","name":"Compas Maped study 160 5 p","costPrice":4109.57,"salePrice":6164.35,"category":"Geometría","stock":0},{"id":"151-256","name":"Compas Maped study con portaminas 05 mm","costPrice":4113.36,"salePrice":6170.04,"category":"Geometría","stock":20},{"id":"151-254","name":"Compas Maped study Pastel estuche","costPrice":3790.55,"salePrice":5685.83,"category":"Geometría","stock":30},{"id":"151-255","name":"Compas Maped study sy 110","costPrice":1986.64,"salePrice":2979.96,"category":"Geometría","stock":15},{"id":"145-127","name":"Compas pizarrón natural","costPrice":11897.98,"salePrice":17846.97,"category":"Geometría","stock":10},{"id":"143-015","name":"Compas Pizzini colegial ARTpk133 caja plástica","costPrice":2755.42,"salePrice":4133.13,"category":"Geometría","stock":5},{"id":"143-016","name":"Compas Pizzini colegial ARTpk133 pote x 12","costPrice":24250.64,"salePrice":36375.96,"category":"Geometría","stock":15},{"id":"143-019","name":"Compas Pizzini plástico c/abrazadera ARTpk120l","costPrice":1656.68,"salePrice":2485.02,"category":"Geometría","stock":5},{"id":"143-007","name":"Compas Pizzini tecnico ARTpk227 metálico circulos 36 cm","costPrice":6658.81,"salePrice":9988.22,"category":"Geometría","stock":10},{"id":"138-092","name":"Perforadora con forma Asb set x 4 tijera formas","costPrice":11683.63,"salePrice":17525.44,"category":"Tijeras","stock":20},{"id":"042-112","name":"Tijera ARTstica 13,5 cm Ibicraft exp x 12 Surtido ART 900801","costPrice":14563.4,"salePrice":21845.1,"category":"Tijeras","stock":0},{"id":"040-987","name":"Tijera escolar Ezco 10 cm mango plástico","costPrice":443.47,"salePrice":598.68,"category":"Tijeras","stock":10},{"id":"040-515","name":"Tijera Ezco crazy stile 14cms con formas","costPrice":1425.13,"salePrice":1923.93,"category":"Tijeras","stock":25},{"id":"040-501","name":"Tijera Ezco Eterna 17cms","costPrice":1561.02,"salePrice":2107.38,"category":"Tijeras","stock":15},{"id":"040-505","name":"Tijera Ezco innova 18cms","costPrice":1822.53,"salePrice":2460.42,"category":"Tijeras","stock":0},{"id":"040-989","name":"Tijera Ezco kids 12cms","costPrice":806.76,"salePrice":1089.12,"category":"Tijeras","stock":0},{"id":"040-518","name":"Tijera Ezco krom metalica 17,5cms blister","costPrice":10351.63,"salePrice":13974.7,"category":"Tijeras","stock":25},{"id":"040-510","name":"Tijera Ezco lotus 22cms","costPrice":2440.35,"salePrice":3294.48,"category":"Tijeras","stock":0},{"id":"040-988","name":"Tijera Ezco pekes 13 cm","costPrice":697.26,"salePrice":941.29,"category":"Tijeras","stock":0},{"id":"040-986","name":"Tijera Ezco zona z 13cms zurdos blister","costPrice":998.1,"salePrice":1347.43,"category":"Tijeras","stock":30},{"id":"034-790","name":"Tijera Filgo escolar craft-me 12 cm Surtido","costPrice":1393.65,"salePrice":2090.48,"category":"Tijeras","stock":5},{"id":"034-093","name":"Tijera Filgo pinto 11,5 cm","costPrice":857.15,"salePrice":1285.72,"category":"Tijeras","stock":30},{"id":"034-092","name":"Tijera Filgo pinto 3D 12 cm","costPrice":1341.85,"salePrice":2012.77,"category":"Tijeras","stock":10},{"id":"042-111","name":"Tijera Ibicraft de precision 15 cm vintage gold ART900654","costPrice":12216.49,"salePrice":18324.74,"category":"Tijeras","stock":0},{"id":"118-114","name":"Adhesivo color Sta amarillo x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":15},{"id":"118-111","name":"Adhesivo color Sta azul x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":0},{"id":"118-117","name":"Adhesivo color Sta blanco x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":5},{"id":"118-087","name":"Adhesivo color Sta celeste x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":15},{"id":"118-113","name":"Adhesivo color Sta marron x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":10},{"id":"118-115","name":"Adhesivo color Sta naranja x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":0},{"id":"118-110","name":"Adhesivo color Sta negro x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":25},{"id":"118-112","name":"Adhesivo color Sta rojo x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":30},{"id":"118-086","name":"Adhesivo color Sta rosa x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":10},{"id":"118-118","name":"Adhesivo color Sta Surtido prim x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":30},{"id":"118-116","name":"Adhesivo color Sta verde x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":0},{"id":"118-119","name":"Adhesivo color Sta violeta x 30 gramos","costPrice":516.94,"salePrice":775.41,"category":"Adhesivos","stock":30},{"id":"183-250","name":"Adhesivo Eccole 9g transparente","costPrice":5040.9,"salePrice":7561.35,"category":"Adhesivos","stock":10},{"id":"165-197","name":"Adhesivo en barra Elmers bloostik buddies x4","costPrice":5199.2,"salePrice":7798.8,"category":"Adhesivos","stock":0},{"id":"165-195","name":"Adhesivo en barra Elmers scented 6 gramos x 4","costPrice":5199.2,"salePrice":7798.8,"category":"Adhesivos","stock":0},{"id":"008-055","name":"Agenda 2026 escritorio con espiral tapa pvc a42l","costPrice":4492.08,"salePrice":6738.12,"category":"Agendas","stock":10},{"id":"005-051","name":"Agenda 2026 Fera 14 x 20 semanal cosida literaria libros","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":0},{"id":"005-060","name":"Agenda 2026 Fera 14 x 20 semanal espiral literaria living de pelota","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":0},{"id":"005-062","name":"Agenda 2026 Fera 14 x 20 semanal espiral tapa dura neutra definiciones","costPrice":18422.25,"salePrice":27633.38,"category":"Agendas","stock":30},{"id":"005-063","name":"Agenda 2026 Fera 14 x 20 semanal espiral tapa dura neutra sabes","costPrice":18422.25,"salePrice":27633.38,"category":"Agendas","stock":0},{"id":"005-031","name":"Agenda 2026 Fera 15 x 21 diaria espiral granadas","costPrice":25282.95,"salePrice":37924.43,"category":"Agendas","stock":10},{"id":"005-030","name":"Agenda 2026 Fera 15 x 21 diaria espiral invernadero","costPrice":25282.95,"salePrice":37924.43,"category":"Agendas","stock":10},{"id":"005-048","name":"Agenda 2026 Fera 15 x 21 semanal espiral astrologica la colgada","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":5},{"id":"005-046","name":"Agenda 2026 Fera 15 x 21 semanal espiral astrologica la emperatriz","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":15},{"id":"005-047","name":"Agenda 2026 Fera 15 x 21 semanal espiral astrologica la fuerza","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":5},{"id":"005-045","name":"Agenda 2026 Fera 15 x 21 semanal espiral astrologicala maga","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":10},{"id":"005-024","name":"Agenda 2026 Fera 15 x 21 semanal espiral corazon","costPrice":20963.25,"salePrice":31444.88,"category":"Agendas","stock":30},{"id":"005-025","name":"Agenda 2026 Fera 15 x 21 semanal espiral estoy bien","costPrice":20963.25,"salePrice":31444.88,"category":"Agendas","stock":25},{"id":"005-021","name":"Agenda 2026 Fera 15 x 21 semanal espiral granadas","costPrice":20963.25,"salePrice":31444.88,"category":"Agendas","stock":10},{"id":"005-020","name":"Agenda 2026 Fera 15 x 21 semanal espiral invernadero","costPrice":20963.25,"salePrice":31444.88,"category":"Agendas","stock":0},{"id":"005-023","name":"Agenda 2026 Fera 15 x 21 semanal espiral la torre","costPrice":20963.25,"salePrice":31444.88,"category":"Agendas","stock":0},{"id":"005-022","name":"Agenda 2026 Fera 15 x 21 semanal espiral mediterraneo","costPrice":20963.25,"salePrice":31444.88,"category":"Agendas","stock":20},{"id":"005-015","name":"Agenda 2026 Fera 15 x 21 semanal espiral naty franz ojo","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":10},{"id":"005-016","name":"Agenda 2026 Fera 15 x 21 semanal espiral naty franz tigre","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":0},{"id":"005-011","name":"Agenda 2026 Fera 15 x 21 semanal espiral paulina cocina bolsa de compras","costPrice":22741.95,"salePrice":34112.93,"category":"Agendas","stock":0},{"id":"166-323","name":"Aprieta papel niquelado Sifap 100 mm c/u","costPrice":1003.16,"salePrice":1504.74,"category":"Papelería","stock":15},{"id":"166-324","name":"Aprieta papel niquelado Sifap 120 mm c/u","costPrice":1442.05,"salePrice":2163.07,"category":"Papelería","stock":0},{"id":"166-326","name":"Aprieta papel niquelado Sifap 145 mm c/u","costPrice":1567.45,"salePrice":2351.18,"category":"Papelería","stock":10},{"id":"166-320","name":"Aprieta papel niquelado Sifap 40 mm c/u","costPrice":410.67,"salePrice":616.0,"category":"Papelería","stock":0},{"id":"166-321","name":"Aprieta papel niquelado Sifap 50 mm c/u","costPrice":478.7,"salePrice":718.05,"category":"Papelería","stock":25},{"id":"166-322","name":"Aprieta papel niquelado Sifap 75 mm c/u","costPrice":661.47,"salePrice":992.21,"category":"Papelería","stock":30},{"id":"143-977","name":"Bandeja papelera Pizzini A4 autoapilable ART15510","costPrice":7586.48,"salePrice":11379.72,"category":"Papelería","stock":20},{"id":"143-974","name":"Bandeja papelera Pizzini oficio autoapilable ART801","costPrice":6970.81,"salePrice":10456.22,"category":"Papelería","stock":30},{"id":"143-971","name":"Bandeja papelera -set separadores -","costPrice":2415.67,"salePrice":3623.51,"category":"Papelería","stock":0},{"id":"143-972","name":"Bandeja papelera spazio oficio apilable -2 pisos","costPrice":24764.19,"salePrice":37146.28,"category":"Papelería","stock":0},{"id":"143-973","name":"Bandeja papelera spazio oficio apilable -3 pisos","costPrice":38535.23,"salePrice":57802.85,"category":"Papelería","stock":0},{"id":"143-952","name":"Bandeja papelera spazio oficio bisagra -2 pisos","costPrice":29599.5,"salePrice":44399.25,"category":"Papelería","stock":0},{"id":"143-953","name":"Bandeja papelera spazio oficio bisagra -3 pisos","costPrice":46059.38,"salePrice":69089.07,"category":"Papelería","stock":30},{"id":"136-089","name":"Bibliorato lomo papel A4 palanca niquelada","costPrice":2676.52,"salePrice":4014.78,"category":"Papelería","stock":0},{"id":"136-088","name":"Bibliorato lomo papel esquela palanca niquelada","costPrice":2676.52,"salePrice":4014.78,"category":"Papelería","stock":15},{"id":"040-982","name":"Acuarela Ezco blister 12 colores + pincel","costPrice":1014.56,"salePrice":1369.65,"category":"Pinceles","stock":30},{"id":"040-980","name":"Acuarela Ezco caja acrilica 12 colores + pincel (base negra)","costPrice":1557.56,"salePrice":2102.7,"category":"Pinceles","stock":0},{"id":"040-983","name":"Acuarela Ezco caja acrilica premium 12 colores + pincel","costPrice":2374.85,"salePrice":3206.04,"category":"Pinceles","stock":20},{"id":"040-984","name":"Acuarela Ezco caja acrilica premium 24 colores + pincel","costPrice":5790.68,"salePrice":7817.41,"category":"Pinceles","stock":0},{"id":"040-985","name":"Acuarela Ezco caja acrilica premium 36 colores + pincel","costPrice":6636.55,"salePrice":8959.35,"category":"Pinceles","stock":0},{"id":"081-050","name":"Acuarela Giotto 12 colores bandeja + pincel","costPrice":4604.58,"salePrice":6906.87,"category":"Pinceles","stock":10},{"id":"081-051","name":"Acuarela Giotto 24 colores bandeja + pincel","costPrice":8717.41,"salePrice":13076.11,"category":"Pinceles","stock":15},{"id":"081-052","name":"Acuarela Giotto 36 colores bandeja + pincel","costPrice":12517.3,"salePrice":18775.95,"category":"Pinceles","stock":0},{"id":"146-230","name":"Acuarela Reeves lata x 12 pastillas + 1 pincel","costPrice":17917.99,"salePrice":26876.99,"category":"Pinceles","stock":15},{"id":"146-232","name":"Acuarela Reeves x 12 pastillas + 1 pincel","costPrice":19979.63,"salePrice":29969.44,"category":"Pinceles","stock":5},{"id":"146-233","name":"Acuarela Reeves x 24 pastillas + 1 pincel de agua","costPrice":33213.68,"salePrice":49820.52,"category":"Pinceles","stock":5},{"id":"069-902","name":"Acuarela Trabi tapa acrilica x12 + pincel","costPrice":2703.07,"salePrice":4054.61,"category":"Pinceles","stock":0},{"id":"069-905","name":"Acuarela Trabi tapa acrilica x28 + pincel","costPrice":5235.05,"salePrice":7852.58,"category":"Pinceles","stock":15},{"id":"069-904","name":"Acuarela Trabi tapa acrilica x48 + pincel","costPrice":8912.86,"salePrice":13369.29,"category":"Pinceles","stock":30},{"id":"082-845","name":"Acuarelas doms 12 colores (23 mm de diametro) + pincel","costPrice":2447.65,"salePrice":3671.48,"category":"Pinceles","stock":20},{"id":"070-910","name":"Acrílico decorativo 200 ml 010-blanco titanio","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"070-911","name":"Acrílico decorativo 200 ml 011-blanco de zinc","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":5},{"id":"070-912","name":"Acrílico decorativo 200 ml 014-blanco antiguo","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"070-932","name":"Acrílico decorativo 200 ml 032-rosa","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"070-935","name":"Acrílico decorativo 200 ml 035-rosa country","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":20},{"id":"070-940","name":"Acrílico decorativo 200 ml 040-amarillo claro","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"070-941","name":"Acrílico decorativo 200 ml 041-amarillo medio","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"070-942","name":"Acrílico decorativo 200 ml 042-amarillo cadmio claro","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":15},{"id":"070-943","name":"Acrílico decorativo 200 ml 043-amarillo cadmio","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":20},{"id":"070-946","name":"Acrílico decorativo 200 ml 046- maiz","costPrice":2831.75,"salePrice":4247.62,"category":"Acrílicos","stock":5},{"id":"070-962","name":"Acrílico decorativo 200 ml 062-naranja permanente","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":10},{"id":"070-966","name":"Acrílico decorativo 200 ml 066-bermellon","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"070-971","name":"Acrílico decorativo 200 ml 071-rojo cadmio","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"070-980","name":"Acrílico decorativo 200 ml 080-rojo Artística dibu AD","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":20},{"id":"070-913","name":"Acrílico decorativo 200 ml 107-capuchino","costPrice":3904.02,"salePrice":5856.03,"category":"Acrílicos","stock":0},{"id":"082-520","name":"Acuarela Aquafine lata 10 colores","costPrice":18344.04,"salePrice":27516.06,"category":"Acuarelas","stock":25},{"id":"082-522","name":"Acuarela Aquafine lata 18 colores","costPrice":25465.84,"salePrice":38198.76,"category":"Acuarelas","stock":5},{"id":"082-524","name":"Acuarela Aquafine lata 24 colores","costPrice":33450.9,"salePrice":50176.35,"category":"Acuarelas","stock":5},{"id":"034-700","name":"Acuarela Filgo pinto estuche x 12 colores","costPrice":1276.25,"salePrice":1914.38,"category":"Acuarelas","stock":25},{"id":"034-701","name":"Acuarela Filgo pinto estuche x 36 colores","costPrice":9641.92,"salePrice":14462.88,"category":"Acuarelas","stock":15},{"id":"146-192","name":"Acuarela Reeves x 12 pomos surtidos de 12 ml","costPrice":11985.43,"salePrice":17978.15,"category":"Acuarelas","stock":10},{"id":"146-193","name":"Acuarela Reeves x 18 pomos surtidos de 12 ml","costPrice":16779.65,"salePrice":25169.48,"category":"Acuarelas","stock":5},{"id":"166-863","name":"Acuarela Sifap pocket x 12 colores apilable","costPrice":1933.35,"salePrice":2900.02,"category":"Acuarelas","stock":5},{"id":"082-504","name":"Acuarela Simply 12 ml set 12 colores","costPrice":14891.05,"salePrice":22336.57,"category":"Acuarelas","stock":0},{"id":"082-505","name":"Acuarela Simply 12 ml set 24 colores","costPrice":28163.49,"salePrice":42245.24,"category":"Acuarelas","stock":5},{"id":"197-003","name":"Abacos 3 columnas","costPrice":3848.53,"salePrice":5772.8,"category":"Otros","stock":20},{"id":"197-004","name":"Abacos 4 columnas","costPrice":5011.77,"salePrice":7517.66,"category":"Otros","stock":15},{"id":"197-005","name":"Abacos 5 columnas","costPrice":6175.02,"salePrice":9262.53,"category":"Otros","stock":10},{"id":"197-006","name":"Abacos 6 columnas","costPrice":7338.27,"salePrice":11007.41,"category":"Otros","stock":0},{"id":"042-500","name":"Abanico Ibicraft helen ART 990850","costPrice":5245.62,"salePrice":7868.43,"category":"Otros","stock":10},{"id":"042-000","name":"Abrochadora 24/6 de mesa Ibicraft 20 hojas ART 20012","costPrice":5179.41,"salePrice":7769.11,"category":"Otros","stock":15},{"id":"044-248","name":"Abrochadora Grap pinza 10/50 p","costPrice":17926.79,"salePrice":26890.19,"category":"Otros","stock":30},{"id":"044-250","name":"Abrochadora Grap pinza 21/6 pintada","costPrice":21158.58,"salePrice":31737.87,"category":"Otros","stock":20},{"id":"044-256","name":"Abrochadora Grap pinza 24/6 pintada","costPrice":20402.92,"salePrice":30604.38,"category":"Otros","stock":30},{"id":"044-365","name":"Abrochadora Grap pinza 65 pintada","costPrice":15729.14,"salePrice":23593.71,"category":"Otros","stock":15}];

const RED = "#c0392b", REDD = "#922b21";
const STAGES = ["reserva","confirmado","en armado","entregado"];
const SCFG = {
  reserva:     {label:"Reserva",    color:"#c0392b", bg:"#fdecea", icon:"🕐"},
  confirmado:  {label:"Confirmado", color:"#1a5276", bg:"#d6eaf8", icon:"✅"},
  "en armado": {label:"En Armado",  color:"#6c3483", bg:"#e8daef", icon:"📦"},
  entregado:   {label:"Entregado",  color:"#1e8449", bg:"#d5f5e3", icon:"🎉"},
};

const LOGO = "/logo.png";

const fARS = n => "$" + Number(n).toLocaleString("es-AR",{minimumFractionDigits:2,maximumFractionDigits:2});
const genId = () => Date.now().toString(36)+Math.random().toString(36).slice(2);
const today = () => new Date().toLocaleDateString("es-AR");

// ─── DEFAULT DATA ────────────────────────────────────────────────────────────
const DEFAULT_USERS = [
  {id:"u1", username:"admin", password:"admin123", role:"admin", name:"Administrador"},
];

// ─── STORAGE HELPERS (memory only) ──────────────────────────────────────────
function useLocalData(key, initial) {
  const [data, setData] = useState(() => {
    try { const s = localStorage.getItem(key); return s ? JSON.parse(s) : initial; } catch { return initial; }
  });
  const setter = (val) => {
    const next = typeof val === "function" ? val(data) : val;
    try { localStorage.setItem(key, JSON.stringify(next)); } catch {}
    setData(next);
  };
  return [data, setter];
}

// ─── SMALL COMPONENTS ────────────────────────────────────────────────────────
function SPill({n}) {
  const [c,b] = n===0?["#c0392b","#fdecea"]:n<=5?["#e67e22","#fef9e7"]:["#1e8449","#d5f5e3"];
  return <span style={{background:b,color:c,borderRadius:10,padding:"2px 8px",fontSize:11,fontWeight:700}}>{n===0?"Sin stock":n<=5?`⚠ ${n}`:n}</span>;
}
function Bdg({stage}) {
  const c=SCFG[stage]||{};
  return <span style={{background:c.bg,color:c.color,border:`1px solid ${c.color}44`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{c.icon} {c.label}</span>;
}
function Field({label,children}) {
  return <div style={{marginBottom:12}}>
    <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>{label}</label>
    {children}
  </div>;
}
const inputStyle = {width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box",background:"#fff"};


// ─── NOTIFICATION TYPES CONFIG ───────────────────────────────────────────────
const NOTIF_TYPES = {
  NUEVO_PEDIDO:   {label:"Nuevo pedido",         icon:"🛒", color:"#1a5276", bg:"#d6eaf8"},
  CAMBIO_ESTADO:  {label:"Cambio de estado",      icon:"📋", color:"#6c3483", bg:"#e8daef"},
  ALTA_MERCADERIA:{label:"Alta de mercadería",    icon:"📦", color:"#1e8449", bg:"#d5f5e3"},
  PEDIDOS_PEND:   {label:"Pedidos pendientes",    icon:"⏰", color:"#e67e22", bg:"#fef9e7"},
};

// ─── NOTIF PANEL (dropdown) ───────────────────────────────────────────────────
function NotifPanel({notifs,setNotifs,currentUser,users,onClose,onMarkAllRead,pushNotif,orders}) {
  const myNotifs = notifs.filter(n =>
    n.para === "todos" || n.para === currentUser.role || n.para === currentUser.id
  );
  const unread = myNotifs.filter(n => !n.leida.includes(currentUser.id));

  const markRead = (id) => setNotifs(ns => ns.map(n =>
    n.id===id && !n.leida.includes(currentUser.id)
      ? {...n, leida:[...n.leida, currentUser.id]}
      : n
  ));
  const delNotif = (id) => setNotifs(ns => ns.filter(n => n.id!==id));

  const pendingOrders = orders.filter(o=>o.stage!=="entregado");

  return (
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:998}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{
        position:"fixed",top:70,right:16,width:380,maxHeight:"80vh",
        background:"#fff",borderRadius:16,boxShadow:"0 12px 40px #0003",
        border:"1px solid #f0f0f0",display:"flex",flexDirection:"column",zIndex:999,overflow:"hidden"
      }}>
        <div style={{padding:"14px 18px",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",justifyContent:"space-between",background:`linear-gradient(135deg,${REDD},${RED})`,borderRadius:"16px 16px 0 0"}}>
          <div style={{color:"#fff",fontWeight:800,fontSize:15}}>
            🔔 Notificaciones {unread.length>0&&<span style={{background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:11,padding:"1px 6px",marginLeft:6,fontWeight:800}}>{unread.length}</span>}
          </div>
          <div style={{display:"flex",gap:6}}>
            {unread.length>0&&<button onClick={onMarkAllRead} style={{padding:"4px 10px",borderRadius:6,border:"none",background:"#ffffff33",color:"#fff",fontSize:11,fontWeight:600,cursor:"pointer"}}>Marcar todas leídas</button>}
            <button onClick={onClose} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"4px 8px",cursor:"pointer",fontSize:14}}>✕</button>
          </div>
        </div>

        {currentUser.role==="vendedor" && pendingOrders.filter(o=>o.vendedor===currentUser.name).length>0&&(
          <div style={{background:"#fef9e7",borderBottom:"1px solid #f1c40f22",padding:"10px 16px",display:"flex",gap:8,alignItems:"center"}}>
            <span style={{fontSize:18}}>⏰</span>
            <div>
              <div style={{fontWeight:700,fontSize:12,color:"#7d6608"}}>Tenés {pendingOrders.filter(o=>o.vendedor===currentUser.name).length} pedido(s) pendiente(s)</div>
              <div style={{fontSize:11,color:"#9a7d0a"}}>Revisá el estado de tus pedidos en Central</div>
            </div>
          </div>
        )}

        <div style={{overflowY:"auto",flex:1}}>
          {myNotifs.length===0
            ? <div style={{textAlign:"center",padding:40,color:"#aaa"}}>
                <div style={{fontSize:36,marginBottom:8}}>🔕</div>
                <div>No tenés notificaciones</div>
              </div>
            : myNotifs.map(n=>{
                const cfg = NOTIF_TYPES[n.tipo]||{icon:"•",color:"#666",bg:"#f5f5f5"};
                const isRead = n.leida.includes(currentUser.id);
                return (
                  <div key={n.id} onClick={()=>markRead(n.id)} style={{
                    padding:"12px 16px",borderBottom:"1px solid #f9f9f9",
                    background:isRead?"#fff":"#fafbff",cursor:"pointer",
                    display:"flex",gap:10,alignItems:"flex-start",
                  }}>
                    <span style={{fontSize:20,flexShrink:0,marginTop:2}}>{cfg.icon}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:6}}>
                        <div style={{fontWeight:isRead?500:700,fontSize:13,color:"#1a1a1a",lineHeight:1.3}}>{n.titulo}</div>
                        {!isRead&&<span style={{width:8,height:8,borderRadius:"50%",background:RED,flexShrink:0,marginTop:4}}/>}
                      </div>
                      <div style={{fontSize:12,color:"#666",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{n.cuerpo}</div>
                      <div style={{fontSize:10,color:"#aaa",marginTop:4}}>{n.fecha}</div>
                    </div>
                    <button onClick={e=>{e.stopPropagation();delNotif(n.id);}} style={{background:"none",border:"none",color:"#ddd",cursor:"pointer",fontSize:16,flexShrink:0,padding:0,lineHeight:1}}>×</button>
                  </div>
                );
              })
          }
        </div>
      </div>
    </div>
  );
}

// ─── NOTIF CONFIG (admin panel section) ──────────────────────────────────────
function NotifConfig({users,setUsers,notifs,setNotifs}) {
  const DEFAULT_PREFS = {
    admin:    {NUEVO_PEDIDO:true,  CAMBIO_ESTADO:true,  ALTA_MERCADERIA:true,  PEDIDOS_PEND:false},
    vendedor: {NUEVO_PEDIDO:false, CAMBIO_ESTADO:true,  ALTA_MERCADERIA:false, PEDIDOS_PEND:true},
  };
  const getPrefs = (u) => u.notifPrefs || DEFAULT_PREFS[u.role] || DEFAULT_PREFS.vendedor;
  const togglePref = (uid, tipo) => {
    setUsers(us => us.map(u => {
      if(u.id!==uid) return u;
      const prefs = getPrefs(u);
      return {...u, notifPrefs:{...prefs, [tipo]:!prefs[tipo]}};
    }));
  };

  return (
    <div>
      <div style={{background:"#d6eaf8",border:"1px solid #aed6f1",borderRadius:12,padding:"12px 16px",marginBottom:16,fontSize:13,color:"#1a5276"}}>
        <strong>🔔 Centro de notificaciones</strong> — Configurá qué alertas recibe cada usuario dentro de la app.<br/>
        <span style={{fontSize:11,marginTop:4,display:"block",color:"#1a5276bb"}}>
          Para envío de emails automáticos conectá <strong>EmailJS</strong> cuando la app esté en producción. Los ajustes que configures acá quedan guardados y se usarán para filtrar qué emails se envían.
        </span>
      </div>

      {users.map(u=>{
        const prefs = getPrefs(u);
        return (
          <div key={u.id} style={{background:"#fff",borderRadius:12,padding:18,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
              <span style={{fontSize:24}}>{u.role==="admin"?"👑":"👤"}</span>
              <div>
                <div style={{fontWeight:800,fontSize:14}}>{u.name}</div>
                <div style={{fontSize:11,color:"#888"}}>
                  @{u.username} · <span style={{color:u.role==="admin"?RED:"#1a5276",fontWeight:600}}>{u.role==="admin"?"Admin":"Vendedor"}</span>
                  {u.email&&<span style={{marginLeft:6,color:"#aaa"}}>· 📧 {u.email}</span>}
                  {!u.email&&<span style={{marginLeft:6,color:"#e67e22",fontSize:10}}>⚠ Sin email configurado</span>}
                </div>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8}}>
              {Object.entries(NOTIF_TYPES).map(([tipo,cfg])=>{
                const active = prefs[tipo]||false;
                return (
                  <div key={tipo} onClick={()=>togglePref(u.id,tipo)} style={{
                    display:"flex",alignItems:"center",gap:10,padding:"10px 14px",
                    borderRadius:10,border:`1.5px solid ${active?cfg.color:"#e5e5e5"}`,
                    background:active?cfg.bg:"#fafafa",cursor:"pointer",
                  }}>
                    <span style={{fontSize:18}}>{cfg.icon}</span>
                    <div style={{flex:1}}>
                      <div style={{fontWeight:600,fontSize:12,color:active?cfg.color:"#666"}}>{cfg.label}</div>
                    </div>
                    <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${active?cfg.color:"#ccc"}`,background:active?cfg.color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      {active&&<span style={{color:"#fff",fontSize:11,fontWeight:800}}>✓</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}

      {notifs.length>0&&(
        <div style={{marginTop:8,textAlign:"right"}}>
          <button onClick={()=>setNotifs([])} style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12,fontWeight:600}}>
            🗑 Limpiar todas las notificaciones ({notifs.length})
          </button>
        </div>
      )}
    </div>
  );
}

// ─── LOGIN SCREEN ─────────────────────────────────────────────────────────────
function Login({users, onLogin}) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [showPass, setShowPass] = useState(false);

  const handleLogin = () => {
    const u = users.find(u => u.username === username.trim() && u.password === password);
    if (u) { setError(""); onLogin(u); }
    else setError("Usuario o contraseña incorrectos");
  };

  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg,${REDD},${RED})`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:20,padding:40,width:"100%",maxWidth:380,boxShadow:"0 20px 60px #0004"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <img src={LOGO} alt="LM" style={{width:90,height:90,borderRadius:"50%",objectFit:"cover",marginBottom:12,boxShadow:"0 4px 16px #0002"}}/>
          <div style={{fontWeight:800,fontSize:22,fontFamily:"Georgia,serif",color:"#1a1a1a"}}>Librería LM</div>
          <div style={{fontSize:12,color:"#aaa",letterSpacing:2,textTransform:"uppercase",marginTop:2}}>Sistema de Gestión</div>
        </div>
        <Field label="Usuario">
          <input value={username} onChange={e=>setUsername(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&handleLogin()}
            placeholder="Tu usuario" style={inputStyle}/>
        </Field>
        <Field label="Contraseña">
          <div style={{position:"relative"}}>
            <input type={showPass?"text":"password"} value={password}
              onChange={e=>setPassword(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&handleLogin()}
              placeholder="Tu contraseña" style={{...inputStyle,paddingRight:40}}/>
            <button onClick={()=>setShowPass(s=>!s)} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#aaa"}}>{showPass?"🙈":"👁"}</button>
          </div>
        </Field>
        {error && <div style={{background:"#fdecea",color:RED,borderRadius:8,padding:"8px 12px",fontSize:13,marginBottom:12,textAlign:"center"}}>{error}</div>}
        <button onClick={handleLogin} style={{width:"100%",padding:12,borderRadius:10,border:"none",background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",fontWeight:800,fontSize:15,cursor:"pointer",marginTop:4}}>
          Ingresar
        </button>
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [users, setUsers]         = useLocalData("lm_users", DEFAULT_USERS);
  const [vendors, setVendors]     = useLocalData("lm_vendors", ["Rocío","Valentina","Lucía","Sofía","Martina"]);
  const [products, setProducts]   = useLocalData("lm_products", CATALOG.map(p=>({...p})));
  const [orders, setOrders]       = useLocalData("lm_orders", []);
  const [stockLog, setStockLog]   = useLocalData("lm_stocklog", []);
  const [notifs, setNotifs]       = useLocalData("lm_notifs", []);

  if (!currentUser) return <Login users={users} onLogin={u=>setCurrentUser(u)}/>;

  return <MainApp
    currentUser={currentUser}
    onLogout={()=>setCurrentUser(null)}
    users={users} setUsers={setUsers}
    vendors={vendors} setVendors={setVendors}
    products={products} setProducts={setProducts}
    orders={orders} setOrders={setOrders}
    stockLog={stockLog} setStockLog={setStockLog}
    notifs={notifs} setNotifs={setNotifs}
  />;
}

// ─── MAIN APP (authenticated) ─────────────────────────────────────────────────
function MainApp({currentUser,onLogout,users,setUsers,vendors,setVendors,products,setProducts,orders,setOrders,stockLog,setStockLog,notifs,setNotifs}) {
  const isAdmin = currentUser.role === "admin";
  const [tab, setTab] = useState("central");
  const [showNotifs, setShowNotifs] = useState(false);

  // Push a notification to specific user ids (or "all" / "admins")
  const pushNotif = (notif) => {
    setNotifs(n => [{
      id: genId(),
      fecha: new Date().toLocaleString("es-AR"),
      leida: [],   // array of userIds that marked it read
      ...notif,
    }, ...n]);
  };

  // Unread count for current user
  const unreadCount = notifs.filter(n =>
    !n.leida.includes(currentUser.id) &&
    (n.para === "todos" || n.para === currentUser.role || n.para === currentUser.id)
  ).length;

  const markAllRead = () => setNotifs(ns => ns.map(n =>
    n.leida.includes(currentUser.id) ? n : {...n, leida:[...n.leida, currentUser.id]}
  ));

  const addLog = (entry) => setStockLog(l => [{
    id: genId(),
    fecha: new Date().toLocaleString("es-AR"),
    usuario: currentUser.name,
    rol: currentUser.role,
    ...entry,
  }, ...l]);

  const addOrder = (order) => {
    setProducts(p => p.map(x => {
      const it = order.items.find(i=>i.pid===x.id);
      return it ? {...x, stock:Math.max(0,x.stock-it.qty)} : x;
    }));
    setOrders(o => [order,...o]);
    pushNotif({
      tipo: "NUEVO_PEDIDO",
      para: "admin",
      icono: "🛒",
      titulo: "Nuevo pedido registrado",
      cuerpo: `${order.client} — ${fARS(order.total)} — Vendedor: ${order.vendedor||"—"}`,
      ref: order.id,
    });
  };
  const setStage = (id,stage) => {
    const ord = orders.find(o=>o.id===id);
    setOrders(o=>o.map(x=>x.id===id?{...x,stage}:x));
    if(ord) {
      const cfg = SCFG[stage]||{};
      // Notify admin always
      pushNotif({
        tipo: "CAMBIO_ESTADO",
        para: "admin",
        icono: cfg.icon||"📋",
        titulo: `Pedido pasó a ${cfg.label}`,
        cuerpo: `${ord.client} — ${fARS(ord.total)}`,
        ref: id,
      });
      // Notify the vendedor who created it (find by name match)
      const vendUser = users.find(u=>u.name===ord.vendedor||u.username===ord.vendedor);
      if(vendUser && vendUser.id !== currentUser.id) {
        pushNotif({
          tipo: "CAMBIO_ESTADO",
          para: vendUser.id,
          icono: cfg.icon||"📋",
          titulo: `Tu pedido pasó a ${cfg.label}`,
          cuerpo: `${ord.client} — ${fARS(ord.total)}`,
          ref: id,
        });
      }
    }
  };
  const delOrder = (id) => {
    const ord = orders.find(o=>o.id===id);
    if(ord && ord.stage!=="entregado") {
      setProducts(p=>p.map(x=>{
        const it=ord.items.find(i=>i.pid===x.id);
        return it?{...x,stock:x.stock+it.qty}:x;
      }));
    }
    setOrders(o=>o.filter(x=>x.id!==id));
  };
  const updProd  = (upd) => setProducts(p=>p.map(x=>x.id===upd.id?upd:x));
  const addStock = (pid,qty,newCost) => {
    const prod = products.find(p=>p.id===pid);
    setProducts(p=>p.map(x=>{
      if(x.id!==pid) return x;
      const u = {...x, stock:x.stock+qty};
      if(newCost){ u.costPrice=newCost; u.salePrice=Math.round(newCost*1.5*100)/100; }
      return u;
    }));
    if(prod) pushNotif({
      tipo: "ALTA_MERCADERIA",
      para: "admin",
      icono: "📦",
      titulo: "Alta de mercadería",
      cuerpo: `${prod.name} — +${qty} unidades${newCost?` — Nuevo costo: ${fARS(newCost)}`:""}`,
      ref: pid,
    });
  };

  const pending = orders.filter(o=>o.stage!=="entregado").length;

  const TABS = [
    {k:"central", label:"Central",           icon:"📋", roles:["admin","vendedor"]},
    {k:"nuevo",   label:"Nuevo Pedido",       icon:"🛒", roles:["admin","vendedor"]},
    {k:"stock",   label:"Stock",              icon:"📦", roles:["admin","vendedor"]},
    {k:"compras", label:"Alta de Mercadería", icon:"🏪", roles:["admin","vendedor"]},
    {k:"admin",   label:"Administración",     icon:"⚙️",  roles:["admin"]},
  ].filter(t=>t.roles.includes(currentUser.role));

  return (
    <div style={{minHeight:"100vh",background:"#f5f5f5",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
      {/* HEADER */}
      <div style={{background:`linear-gradient(135deg,${REDD},${RED})`,boxShadow:"0 4px 16px #0004"}}>
        <div style={{maxWidth:1200,margin:"0 auto",padding:"0 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap"}}>
          <div style={{display:"flex",alignItems:"center",gap:12,padding:"12px 0"}}>
            <img src={LOGO} alt="LM Logo" style={{width:54,height:54,borderRadius:"50%",objectFit:"cover",boxShadow:"0 2px 8px #0003"}}/>
            <div>
              <div style={{color:"#fff",fontWeight:800,fontSize:20,fontFamily:"Georgia,serif"}}>Librería LM</div>
              <div style={{color:"#ffcccc",fontSize:10,letterSpacing:2,textTransform:"uppercase"}}>Sistema de Gestión</div>
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:4,flexWrap:"wrap"}}>
            <nav style={{display:"flex",gap:2,flexWrap:"wrap"}}>
              {TABS.map(t=>(
                <button key={t.k} onClick={()=>setTab(t.k)} style={{
                  padding:"11px 14px",border:"none",cursor:"pointer",fontSize:13,
                  background:tab===t.k?"#fff":"transparent",
                  color:tab===t.k?RED:"#ffcccc",
                  fontWeight:tab===t.k?700:500,borderRadius:"8px 8px 0 0",position:"relative",
                }}>
                  {t.icon} {t.label}
                  {t.k==="central"&&pending>0&&<span style={{position:"absolute",top:5,right:3,background:"#fff",color:RED,borderRadius:10,fontSize:10,padding:"1px 5px",fontWeight:800,border:`1.5px solid ${RED}`}}>{pending}</span>}
                </button>
              ))}
            </nav>
            <div style={{display:"flex",alignItems:"center",gap:8,padding:"0 8px",borderLeft:"1px solid #ffffff33",marginLeft:4}}>
              <span style={{color:"#ffeeee",fontSize:12}}>👤 {currentUser.name}</span>
              {/* Bell */}
              <div style={{position:"relative"}}>
                <button onClick={()=>setShowNotifs(s=>!s)} style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"5px 8px",cursor:"pointer",fontSize:16,lineHeight:1,position:"relative"}}>
                  🔔
                  {unreadCount>0&&<span style={{position:"absolute",top:-4,right:-4,background:"#f1c40f",color:"#1a1a1a",borderRadius:10,fontSize:9,padding:"1px 4px",fontWeight:800,minWidth:14,textAlign:"center"}}>{unreadCount}</span>}
                </button>
              </div>
              <button onClick={onLogout} title="Cerrar sesión" style={{background:"#ffffff22",border:"none",color:"#fff",borderRadius:6,padding:"5px 8px",cursor:"pointer",fontSize:11,fontWeight:600}}>Salir</button>
            </div>
            {/* Notif panel dropdown */}
            {showNotifs && <NotifPanel
              notifs={notifs} setNotifs={setNotifs}
              currentUser={currentUser} users={users}
              onClose={()=>setShowNotifs(false)}
              onMarkAllRead={markAllRead}
              pushNotif={pushNotif}
              orders={orders}
            />}
          </div>
        </div>
      </div>

      {/* CONTENT */}
      <div style={{maxWidth:1200,margin:"0 auto",padding:"20px 16px"}}>
        {tab==="central" && <Central orders={orders} products={products} onStage={setStage} onDel={delOrder}/>}
        {tab==="nuevo"   && <Nuevo products={products} vendors={vendors} onAdd={addOrder} onDone={()=>setTab("central")}/>}
        {tab==="stock"   && <Stock products={products} onUpd={updProd}
        onDel={pid=>setProducts(p=>p.filter(x=>x.id!==pid))}
        onAdjust={(pid,qty)=>setProducts(p=>p.map(x=>x.id===pid?{...x,stock:x.stock+qty}:x))}
        isAdmin={isAdmin} addLog={addLog} stockLog={stockLog} setStockLog={setStockLog}/>}
        {tab==="compras" && <Compras products={products} onStock={addStock}/>}
        {tab==="admin"   && isAdmin && <AdminPanel users={users} setUsers={setUsers} vendors={vendors} setVendors={setVendors} products={products} setProducts={setProducts} stockLog={stockLog} setStockLog={setStockLog} notifs={notifs} setNotifs={setNotifs}/>}
      </div>
    </div>
  );
}

// ─── CENTRAL ──────────────────────────────────────────────────────────────────
function Central({orders,products,onStage,onDel}) {
  const [fStage,setFStage]=useState("todos");
  const [search,setSearch]=useState("");
  const [expanded,setExpanded]=useState(null);
  const getP = id=>products.find(p=>p.id===id);
  const filtered = orders.filter(o=>{
    if(fStage!=="todos"&&o.stage!==fStage) return false;
    if(search&&!o.client.toLowerCase().includes(search.toLowerCase())&&!o.id.includes(search)) return false;
    return true;
  });
  const deliv = orders.filter(o=>o.stage==="entregado").reduce((s,o)=>s+o.total,0);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(135px,1fr))",gap:12,marginBottom:20}}>
        {STAGES.map(s=>{
          const c=SCFG[s], cnt=orders.filter(o=>o.stage===s).length;
          return <div key={s} onClick={()=>setFStage(fStage===s?"todos":s)} style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${c.color}`,cursor:"pointer",outline:fStage===s?`2px solid ${c.color}`:"none"}}>
            <div style={{fontSize:26,fontWeight:800,color:c.color}}>{cnt}</div>
            <div style={{fontSize:12,color:"#666",fontWeight:600}}>{c.icon} {c.label}</div>
          </div>;
        })}
        <div style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px #0001",borderLeft:`4px solid ${RED}`}}>
          <div style={{fontSize:14,fontWeight:800,color:RED}}>{fARS(deliv)}</div>
          <div style={{fontSize:12,color:"#666",fontWeight:600}}>💰 Entregado</div>
        </div>
      </div>
      <div style={{background:"#fff",borderRadius:12,padding:14,marginBottom:14,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar cliente o N° pedido..."
          style={{flex:1,minWidth:180,padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {["todos",...STAGES].map(s=>{
            const c=SCFG[s];
            return <button key={s} onClick={()=>setFStage(s)} style={{padding:"5px 11px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:fStage===s?(c?.color||RED):"#e5e5e5",background:fStage===s?(c?.bg||"#fdecea"):"#fff",color:fStage===s?(c?.color||RED):"#666"}}>{s==="todos"?"Todos":c.label}</button>;
          })}
        </div>
      </div>
      {filtered.length===0
        ? <div style={{textAlign:"center",padding:60,color:"#aaa"}}><div style={{fontSize:48}}>📭</div><div style={{marginTop:8}}>No hay pedidos. ¡Creá uno desde "Nuevo Pedido"!</div></div>
        : filtered.map(o=><OCard key={o.id} o={o} exp={expanded===o.id} toggle={()=>setExpanded(expanded===o.id?null:o.id)} getP={getP} onStage={onStage} onDel={onDel}/>)
      }
    </div>
  );
}


function DelBtn({onConfirm}) {
  const [confirm, setConfirm] = useState(false);
  if(confirm) return (
    <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:6,background:"#fdecea",borderRadius:8,padding:"6px 10px",border:"1.5px solid #fcc"}}>
      <span style={{fontSize:12,color:RED,fontWeight:600}}>¿Eliminar?</span>
      <button onClick={onConfirm} style={{padding:"4px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>Sí</button>
      <button onClick={()=>setConfirm(false)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
    </div>
  );
  return <button onClick={()=>setConfirm(true)} style={{marginLeft:"auto",padding:"8px 12px",borderRadius:8,border:"1.5px solid #fcc",cursor:"pointer",background:"#fff",color:RED,fontWeight:600,fontSize:13}}>🗑 Eliminar</button>;
}

function OCard({o,exp,toggle,getP,onStage,onDel}) {
  const idx=STAGES.indexOf(o.stage), next=STAGES[idx+1];
  return (
    <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 6px #0001",overflow:"hidden",marginBottom:8}}>
      <div onClick={toggle} style={{padding:"13px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer"}}>
        <div style={{flex:1,minWidth:100}}>
          <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{o.client}</div>
          <div style={{fontSize:11,color:"#aaa"}}>#{o.id.slice(-6).toUpperCase()} · {o.date}{o.vendedor?<span style={{marginLeft:6}}>· 👤 {o.vendedor}</span>:null}</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <Bdg stage={o.stage}/>
          <span style={{fontWeight:800,color:RED,fontSize:15}}>{fARS(o.total)}</span>
          <span style={{color:"#ccc"}}>{exp?"▲":"▼"}</span>
        </div>
      </div>
      {exp && (
        <div style={{borderTop:"1px solid #f5f5f5",padding:18}}>
          <div style={{display:"flex",marginBottom:18,overflowX:"auto"}}>
            {STAGES.map((s,i)=>{
              const done=i<=idx, c=SCFG[s];
              return <div key={s} style={{display:"flex",alignItems:"center",flex:1,minWidth:65}}>
                <div style={{textAlign:"center",flex:1}}>
                  <div style={{width:30,height:30,borderRadius:"50%",background:done?c.color:"#eee",color:done?"#fff":"#aaa",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 4px",fontSize:13}}>{done?c.icon:"○"}</div>
                  <div style={{fontSize:10,color:done?c.color:"#aaa",fontWeight:done?700:400}}>{c.label}</div>
                </div>
                {i<3&&<div style={{height:2,background:i<idx?RED:"#eee",flex:1}}/>}
              </div>;
            })}
          </div>
          {o.items.map((it,i)=>{
            const p=getP(it.pid);
            return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f9f9f9",fontSize:13}}>
              <span style={{color:"#444"}}>{p?.name||it.name} × {it.qty}</span>
              <span style={{fontWeight:600}}>{fARS(it.price*it.qty)}</span>
            </div>;
          })}
          <div style={{display:"flex",justifyContent:"flex-end",fontWeight:800,fontSize:16,color:RED,margin:"8px 0 12px"}}>{fARS(o.total)}</div>
          {o.notes&&<div style={{background:"#f9f9f9",borderRadius:8,padding:"8px 12px",fontSize:13,color:"#555",marginBottom:12}}>💬 {o.notes}</div>}
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {next&&<button onClick={()=>onStage(o.id,next)} style={{padding:"8px 14px",borderRadius:8,border:"none",cursor:"pointer",background:SCFG[next].color,color:"#fff",fontWeight:700,fontSize:13}}>{SCFG[next].icon} Pasar a {SCFG[next].label}</button>}
            {idx>0&&o.stage!=="entregado"&&<button onClick={()=>onStage(o.id,STAGES[idx-1])} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",cursor:"pointer",background:"#fff",color:"#666",fontWeight:600,fontSize:13}}>← Retroceder</button>}
            <DelBtn onConfirm={()=>onDel(o.id)}/>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── NUEVO PEDIDO ─────────────────────────────────────────────────────────────
function Nuevo({products,vendors,onAdd,onDone}) {
  const [client,setClient]=useState("");
  const [notes,setNotes]=useState("");
  const [vendedor,setVendedor]=useState("");
  const [cart,setCart]=useState([]);
  const [search,setSearch]=useState("");
  const [cat,setCat]=useState("todos");
  const [catOpen,setCatOpen]=useState(false);
  const [ok,setOk]=useState(false);
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const shown = useMemo(()=>{
    const q=search.toLowerCase();
    return products.filter(p=>{
      if(cat!=="todos"&&p.category!==cat) return false;
      if(q) return p.name.toLowerCase().includes(q)||p.id.includes(q);
      return true;
    }).slice(0,80);
  },[products,search,cat]);
  const addC = p=>setCart(c=>{const ex=c.find(i=>i.pid===p.id);return ex?c.map(i=>i.pid===p.id?{...i,qty:i.qty+1}:i):[...c,{pid:p.id,qty:1,price:p.salePrice,name:p.name}];});
  const setQ = (pid,qty)=>{if(qty<=0)setCart(c=>c.filter(i=>i.pid!==pid));else setCart(c=>c.map(i=>i.pid===pid?{...i,qty}:i));};
  const total=cart.reduce((s,i)=>s+i.price*i.qty,0);
  const submit=()=>{
    if(!client.trim()){alert("Ingresá el cliente");return;}
    if(!vendedor){alert("Seleccioná un vendedor");return;}
    if(!cart.length){alert("Agregá productos");return;}
    onAdd({id:genId(),client:client.trim(),notes,vendedor,items:cart,total,stage:"reserva",date:today()});
    setOk(true); setTimeout(()=>onDone(),1400);
  };
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>✅</div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>¡Pedido registrado!</div></div>;
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 330px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>🛒 Nuevo Pedido — Seleccioná productos</div>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar por nombre o código..."
            style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",marginBottom:10,boxSizing:"border-box"}}/>
          <div style={{position:"relative",marginBottom:search?4:0}}>
            <button onClick={()=>setCatOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${catOpen?RED:"#e5e5e5"}`,background:cat!=="todos"?"#fdecea":"#fff",color:cat!=="todos"?RED:"#666",cursor:"pointer",fontSize:13,fontWeight:600}}>
              <span>🏷️ {cat==="todos"?"Todas las categorías":cat}</span>
              <span style={{fontSize:10,marginLeft:6}}>{catOpen?"▲":"▼"}</span>
            </button>
            {catOpen&&(
              <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",borderRadius:10,border:"1.5px solid #e5e5e5",boxShadow:"0 8px 24px #0002",zIndex:50,padding:8,display:"flex",flexWrap:"wrap",gap:5,maxHeight:220,overflowY:"auto"}}>
                {CATS.map(c=><button key={c} onClick={()=>{setCat(c);setCatOpen(false);}} style={{padding:"4px 11px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:cat===c?RED:"#e5e5e5",background:cat===c?"#fdecea":"#fff",color:cat===c?RED:"#666"}}>{c==="todos"?"Todos":c}</button>)}
              </div>
            )}
          </div>
          {search&&<div style={{fontSize:11,color:"#aaa",marginTop:6}}>{shown.length} resultados</div>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
          {shown.map(p=>{
            const ic=cart.find(i=>i.pid===p.id);
            return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:ic?`2px solid ${RED}`:"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
              <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
              <div style={{fontSize:12,color:"#666",marginBottom:7,fontWeight:500}}>{p.id} · {p.category}</div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <span style={{fontSize:17,fontWeight:800,color:RED}}>{fARS(p.salePrice)}</span>
                <SPill n={p.stock}/>
              </div>
              {ic
                ? <div style={{display:"flex",alignItems:"center",gap:5}}>
                    <button onClick={()=>setQ(p.id,ic.qty-1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>−</button>
                    <input type="number" value={ic.qty} onChange={e=>setQ(p.id,+e.target.value||0)} style={{width:40,textAlign:"center",padding:3,borderRadius:6,border:`1.5px solid ${RED}`,fontWeight:700,fontSize:13,outline:"none"}}/>
                    <button onClick={()=>setQ(p.id,ic.qty+1)} style={{width:27,height:27,borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:15,fontWeight:700}}>+</button>
                    <span style={{color:"#1e8449",fontSize:12,fontWeight:700}}>✓</span>
                  </div>
                : <button onClick={()=>addC(p)} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",cursor:"pointer",background:RED,color:"#fff",fontWeight:700,fontSize:12}}>+ Agregar</button>
              }
            </div>;
          })}
        </div>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>📋 Resumen</div>
          <Field label="Cliente *">
            <input value={client} onChange={e=>setClient(e.target.value)} placeholder="Nombre del cliente" style={inputStyle}/>
          </Field>
          <Field label="Vendedor *">
            <select value={vendedor} onChange={e=>setVendedor(e.target.value)} style={{...inputStyle,color:vendedor?"#1a1a1a":"#aaa",cursor:"pointer"}}>
              <option value="">— Seleccioná vendedor —</option>
              {vendors.map(v=><option key={v} value={v}>{v}</option>)}
            </select>
          </Field>
          <Field label="Notas">
            <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Observaciones..."
              style={{...inputStyle,resize:"vertical",minHeight:55,fontSize:12}}/>
          </Field>
          <div style={{borderTop:"1px solid #f5f5f5",margin:"4px 0 10px",paddingTop:10}}>
            {cart.length===0
              ? <div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"10px 0"}}>Agregá productos al pedido</div>
              : cart.map(i=><div key={i.pid} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"3px 0",color:"#555"}}>
                  <span style={{flex:1,marginRight:6,lineHeight:1.3}}>{i.name} × {i.qty}</span>
                  <span style={{fontWeight:600,whiteSpace:"nowrap"}}>{fARS(i.price*i.qty)}</span>
                </div>)
            }
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:17,color:RED,padding:"8px 0",borderTop:"2px solid #f5f5f5",marginBottom:14}}>
            <span>Total</span><span>{fARS(total)}</span>
          </div>
          <button onClick={submit} disabled={!cart.length||!client.trim()||!vendedor} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:(!cart.length||!client.trim()||!vendedor)?"#e5e5e5":`linear-gradient(135deg,${REDD},${RED})`,color:(!cart.length||!client.trim()||!vendedor)?"#aaa":"#fff"}}>
            ✅ Registrar como Reserva
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── STOCK ────────────────────────────────────────────────────────────────────

function StockAlert({low}) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{background:"#fef9e7",border:"1px solid #f1c40f",borderRadius:12,marginBottom:14,overflow:"hidden"}}>
      <div onClick={()=>setOpen(o=>!o)} style={{padding:"11px 16px",display:"flex",gap:10,alignItems:"center",cursor:"pointer",userSelect:"none"}}>
        <span style={{fontSize:20}}>⚠️</span>
        <div style={{flex:1}}>
          <span style={{fontWeight:700,color:"#7d6608",fontSize:13}}>Stock bajo en {low.length} producto(s)</span>
          {!open && <span style={{fontSize:11,color:"#9a7d0a",marginLeft:8}}>— hacé clic para ver el detalle</span>}
        </div>
        <span style={{fontSize:12,color:"#9a7d0a",fontWeight:700}}>{open?"▲":"▼"}</span>
      </div>
      {open && (
        <div style={{borderTop:"1px solid #f1c40f22",padding:"4px 16px 14px"}}>
          {low.map(p=>(
            <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #f1c40f22"}}>
              <div>
                <span style={{fontWeight:600,fontSize:13,color:"#555"}}>{p.name}</span>
                <span style={{fontSize:11,color:"#aaa",marginLeft:8}}>{p.id}</span>
              </div>
              <span style={{background:p.stock===0?"#fdecea":"#fff3cd",color:p.stock===0?RED:"#856404",borderRadius:8,padding:"2px 8px",fontSize:12,fontWeight:700,border:`1px solid ${p.stock===0?"#f5c6cb":"#ffc107"}`}}>
                {p.stock===0?"Sin stock":`⚠ ${p.stock} u.`}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}



function StockLog({log, onClear}) {
  const [filter,setFilter] = useState("todos");
  const [search,setSearch] = useState("");

  const TIPOS = {
    "ENTRADA": {color:"#1e8449", bg:"#d5f5e3", icon:"📥"},
    "SALIDA":  {color:"#e67e22", bg:"#fef9e7", icon:"📤"},
    "BAJA":    {color:"#c0392b", bg:"#fdecea", icon:"🗑"},
  };

  const filtered = log.filter(e=>{
    if(filter!=="todos" && e.tipo!==filter) return false;
    if(search){
      const q=search.toLowerCase();
      return e.producto?.toLowerCase().includes(q)||e.usuario?.toLowerCase().includes(q)||e.motivo?.toLowerCase().includes(q)||e.productoId?.toLowerCase().includes(q);
    }
    return true;
  });

  const [confirmClear, setConfirmClear] = useState(false);

  return (
    <div>
      {/* Resumen */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10,marginBottom:14}}>
        {Object.entries(TIPOS).map(([tipo,cfg])=>{
          const cnt = log.filter(e=>e.tipo===tipo).length;
          return <div key={tipo} style={{background:"#fff",borderRadius:10,padding:"12px 14px",boxShadow:"0 1px 4px #0001",borderLeft:`4px solid ${cfg.color}`}}>
            <div style={{fontSize:22,fontWeight:800,color:cfg.color}}>{cnt}</div>
            <div style={{fontSize:12,color:"#666",fontWeight:600}}>{cfg.icon} {tipo === "ENTRADA" ? "Entradas" : tipo === "SALIDA" ? "Salidas" : "Bajas"}</div>
          </div>;
        })}
        <div style={{background:"#fff",borderRadius:10,padding:"12px 14px",boxShadow:"0 1px 4px #0001",borderLeft:"4px solid #1a5276"}}>
          <div style={{fontSize:22,fontWeight:800,color:"#1a5276"}}>{log.length}</div>
          <div style={{fontSize:12,color:"#666",fontWeight:600}}>📜 Total movimientos</div>
        </div>
      </div>

      {/* Filtros */}
      <div style={{background:"#fff",borderRadius:12,padding:12,marginBottom:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar por producto, usuario o motivo..."
          style={{flex:1,minWidth:200,padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none"}}/>
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {["todos","ENTRADA","SALIDA","BAJA"].map(t=>{
            const cfg = TIPOS[t];
            return <button key={t} onClick={()=>setFilter(t)} style={{padding:"5px 12px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:filter===t?(cfg?.color||RED):"#e5e5e5",background:filter===t?(cfg?.bg||"#fdecea"):"#fff",color:filter===t?(cfg?.color||RED):"#666"}}>
              {t==="todos"?"Todos":`${cfg.icon} ${t}`}
            </button>;
          })}
        </div>
        {log.length>0&&(
          confirmClear
            ? <div style={{display:"flex",gap:6,alignItems:"center",background:"#fdecea",borderRadius:8,padding:"5px 10px",border:`1px solid ${RED}44`}}>
                <span style={{fontSize:12,color:RED,fontWeight:600}}>¿Limpiar todo el historial?</span>
                <button onClick={()=>{onClear();setConfirmClear(false);}} style={{padding:"3px 10px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>Sí</button>
                <button onClick={()=>setConfirmClear(false)} style={{padding:"3px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>No</button>
              </div>
            : <button onClick={()=>setConfirmClear(true)} style={{padding:"5px 12px",borderRadius:8,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>🗑 Limpiar historial</button>
        )}
      </div>

      {/* Tabla */}
      {filtered.length===0
        ? <div style={{textAlign:"center",padding:50,color:"#aaa",background:"#fff",borderRadius:12}}>
            <div style={{fontSize:40,marginBottom:8}}>📭</div>
            <div>{log.length===0?"No hay movimientos registrados aún.":"No hay movimientos que coincidan con el filtro."}</div>
          </div>
        : <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f9f9f9"}}>
                  {["Fecha","Tipo","Producto","Código","Antes","Cambio","Después","Usuario","Motivo"].map(h=>
                    <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.4,whiteSpace:"nowrap"}}>{h}</th>
                  )}
                </tr>
              </thead>
              <tbody>
                {filtered.map(e=>{
                  const cfg = TIPOS[e.tipo]||{color:"#666",bg:"#f5f5f5",icon:"•"};
                  return <tr key={e.id} style={{borderTop:"1px solid #f5f5f5"}}>
                    <td style={{padding:"9px 12px",fontSize:11,color:"#888",whiteSpace:"nowrap"}}>{e.fecha}</td>
                    <td style={{padding:"9px 12px"}}>
                      <span style={{background:cfg.bg,color:cfg.color,borderRadius:8,padding:"2px 8px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                        {cfg.icon} {e.tipo}
                      </span>
                    </td>
                    <td style={{padding:"9px 12px",fontWeight:600,color:"#1a1a1a",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={e.producto}>{e.producto}</td>
                    <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{e.productoId}</td>
                    <td style={{padding:"9px 12px",color:"#666",textAlign:"center"}}>{e.stockAntes}</td>
                    <td style={{padding:"9px 12px",textAlign:"center",fontWeight:800,color:e.cambio>0?"#1e8449":RED}}>
                      {e.cambio>0?"+":""}{e.cambio}
                    </td>
                    <td style={{padding:"9px 12px",textAlign:"center",fontWeight:700,color:e.stockDespues===0?RED:"#1a1a1a"}}>{e.stockDespues}</td>
                    <td style={{padding:"9px 12px",whiteSpace:"nowrap"}}>
                      <span style={{fontWeight:600,fontSize:12}}>{e.usuario}</span>
                      <span style={{fontSize:10,color:e.rol==="admin"?RED:"#1a5276",marginLeft:4,fontWeight:600}}>({e.rol})</span>
                    </td>
                    <td style={{padding:"9px 12px",color:"#555",fontSize:12,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={e.motivo}>{e.motivo}</td>
                  </tr>;
                })}
              </tbody>
            </table>
          </div>
      }
    </div>
  );
}

function StockAdjust({products,onDel,onAdjust,addLog}) {
  const [search,setSearch]=useState("");
  const [selected,setSelected]=useState(null);
  const [mode,setMode]=useState("ajuste"); // "ajuste" | "baja"
  const [qty,setQty]=useState(0);
  const [reason,setReason]=useState("");
  const [done,setDone]=useState(null);

  const q=search.toLowerCase();
  const found = useMemo(()=>
    q.length>1 ? products.filter(p=>p.name.toLowerCase().includes(q)||p.id.includes(q)).slice(0,40) : []
  ,[products,search,q]);

  const reset = () => { setSelected(null); setQty(0); setReason(""); setSearch(""); };

  const confirmAction = () => {
    if(!selected) return;
    if(!reason.trim()){alert("El motivo es obligatorio para registrar el movimiento");return;}
    if(mode==="baja") {
      onDel(selected.id);
      addLog({
        tipo: "BAJA",
        productoId:   selected.id,
        producto:     selected.name,
        stockAntes:   selected.stock,
        stockDespues: 0,
        cambio:       -selected.stock,
        motivo:       reason.trim(),
      });
      setDone({msg:`✅ Producto "${selected.name}" dado de baja.`, color:"#c0392b"});
    } else {
      if(qty===0){alert("El ajuste no puede ser 0");return;}
      const stockDespues = Math.max(0, selected.stock + qty);
      onAdjust(selected.id, qty);
      addLog({
        tipo:         qty > 0 ? "ENTRADA" : "SALIDA",
        productoId:   selected.id,
        producto:     selected.name,
        stockAntes:   selected.stock,
        stockDespues,
        cambio:       qty,
        motivo:       reason.trim(),
      });
      const signo = qty>0?"+":"";
      setDone({msg:`✅ Stock de "${selected.name}" ajustado en ${signo}${qty} unidades. Nuevo stock: ${stockDespues}`, color:"#1e8449"});
    }
    reset();
    setTimeout(()=>setDone(null),5000);
  };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      {/* Búsqueda */}
      <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>🔍 Buscar producto</div>
        <input value={search} onChange={e=>{setSearch(e.target.value);setSelected(null);}}
          placeholder="Nombre o código del producto..."
          style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:10}}/>
        {done&&<div style={{background:done.color==="#c0392b"?"#fdecea":"#d5f5e3",color:done.color,borderRadius:8,padding:"9px 12px",fontSize:13,fontWeight:600,marginBottom:10}}>{done.msg}</div>}
        <div style={{maxHeight:380,overflowY:"auto"}}>
          {search.length<=1
            ? <div style={{color:"#aaa",fontSize:13,textAlign:"center",padding:20}}>Escribí al menos 2 caracteres para buscar</div>
            : found.length===0
              ? <div style={{color:"#aaa",fontSize:13,textAlign:"center",padding:20}}>Sin resultados</div>
              : found.map(p=>(
                  <div key={p.id} onClick={()=>setSelected(p)}
                    style={{padding:"10px 12px",borderRadius:10,border:`1.5px solid ${selected?.id===p.id?RED:"#f0f0f0"}`,background:selected?.id===p.id?"#fdecea":"#fafafa",cursor:"pointer",marginBottom:6}}>
                    <div style={{fontWeight:600,fontSize:13,color:"#1a1a1a"}}>{p.name}</div>
                    <div style={{fontSize:11,color:"#888",marginTop:2,display:"flex",gap:12}}>
                      <span>{p.id}</span>
                      <span>Stock: <strong style={{color:p.stock===0?RED:p.stock<=5?"#e67e22":"#1e8449"}}>{p.stock===0?"Sin stock":p.stock}</strong></span>
                      <span style={{color:"#aaa"}}>{p.category}</span>
                    </div>
                  </div>
                ))
          }
        </div>
      </div>

      {/* Acción */}
      <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>⚙️ Acción sobre el producto</div>
        {!selected
          ? <div style={{color:"#aaa",fontSize:13,textAlign:"center",padding:30}}>Seleccioná un producto de la lista</div>
          : <>
              {/* Producto seleccionado */}
              <div style={{background:"#f9f9f9",borderRadius:10,padding:"12px 14px",marginBottom:16}}>
                <div style={{fontWeight:700,fontSize:14,color:"#1a1a1a"}}>{selected.name}</div>
                <div style={{fontSize:12,color:"#888",marginTop:3,display:"flex",gap:16}}>
                  <span>{selected.id}</span>
                  <span>Stock actual: <strong style={{color:selected.stock===0?RED:"#1e8449"}}>{selected.stock}</strong></span>
                  <span>Venta: <strong style={{color:RED}}>{fARS(selected.salePrice)}</strong></span>
                </div>
              </div>

              {/* Modo */}
              <div style={{display:"flex",gap:8,marginBottom:16}}>
                {[{v:"ajuste",l:"⚖️ Ajustar stock",c:"#1a5276",bg:"#d6eaf8"},{v:"baja",l:"🗑 Dar de baja",c:RED,bg:"#fdecea"}].map(opt=>(
                  <button key={opt.v} onClick={()=>setMode(opt.v)} style={{flex:1,padding:"9px",borderRadius:9,border:`2px solid ${mode===opt.v?opt.c:"#e5e5e5"}`,background:mode===opt.v?opt.bg:"#fff",color:mode===opt.v?opt.c:"#666",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                    {opt.l}
                  </button>
                ))}
              </div>

              {mode==="ajuste" && <>
                <div style={{marginBottom:12}}>
                  <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>
                    Cantidad a ajustar <span style={{color:"#aaa",fontWeight:400}}>(positivo para agregar, negativo para restar)</span>
                  </label>
                  <input type="number" value={qty} onChange={e=>setQty(+e.target.value||0)}
                    style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:14,outline:"none",boxSizing:"border-box",textAlign:"center",fontWeight:700}}/>
                  {qty!==0&&<div style={{fontSize:12,color:qty>0?"#1e8449":RED,marginTop:6,fontWeight:600,textAlign:"center"}}>
                    Stock resultante: {Math.max(0, selected.stock + qty)} unidades
                  </div>}
                </div>
                <div style={{marginBottom:14}}>
                  <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>Motivo <span style={{color:RED}}>*</span></label>
                  <input value={reason} onChange={e=>setReason(e.target.value)} placeholder="Ej: corrección de inventario, mercadería dañada..."
                    style={{width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${!reason.trim()?"#e5e5e5":"#1e8449"}`,fontSize:13,outline:"none",boxSizing:"border-box"}}/>
                  <div style={{fontSize:11,color:"#aaa",marginTop:3}}>Obligatorio — quedará registrado en el historial de movimientos</div>
                </div>
                <button onClick={confirmAction} disabled={qty===0||!reason.trim()} style={{width:"100%",padding:"10px",borderRadius:9,border:"none",background:(qty===0||!reason.trim())?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",color:(qty===0||!reason.trim())?"#aaa":"#fff",fontWeight:800,fontSize:14,cursor:(qty===0||!reason.trim())?"not-allowed":"pointer"}}>
                  ✅ Confirmar ajuste
                </button>
              </>}

              {mode==="baja" && <>
                <div style={{background:"#fdecea",border:"1px solid #f5c6cb",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
                  <div style={{fontWeight:700,color:RED,fontSize:13,marginBottom:4}}>⚠️ Dar de baja elimina el producto del catálogo</div>
                  <div style={{fontSize:12,color:"#666"}}>Esta acción no se puede deshacer. El producto dejará de aparecer en Nuevo Pedido y Stock.</div>
                </div>
                <div style={{marginBottom:14}}>
                  <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>Motivo de la baja *</label>
                  <input value={reason} onChange={e=>setReason(e.target.value)} placeholder="Ej: producto descontinuado, error de carga..."
                    style={{width:"100%",padding:"8px 12px",borderRadius:8,border:`1.5px solid ${reason?"#e5e5e5":RED+"88"}`,fontSize:13,outline:"none",boxSizing:"border-box"}}/>
                </div>
                <BajaConfirm onConfirm={confirmAction} disabled={!reason.trim()}/>
              </>}
            </>
        }
      </div>
    </div>
  );
}

function BajaConfirm({onConfirm,disabled}) {
  const [step,setStep]=useState(false);
  if(step) return (
    <div style={{background:"#fdecea",borderRadius:9,padding:"12px 14px",border:`1.5px solid ${RED}44`}}>
      <div style={{fontWeight:700,color:RED,fontSize:13,marginBottom:10,textAlign:"center"}}>¿Confirmás la baja definitiva?</div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={onConfirm} style={{flex:1,padding:"9px",borderRadius:8,border:"none",background:RED,color:"#fff",fontWeight:800,cursor:"pointer",fontSize:13}}>Sí, dar de baja</button>
        <button onClick={()=>setStep(false)} style={{flex:1,padding:"9px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666",fontSize:13}}>Cancelar</button>
      </div>
    </div>
  );
  return (
    <button onClick={()=>setStep(true)} disabled={disabled} style={{width:"100%",padding:"10px",borderRadius:9,border:"none",background:disabled?"#e5e5e5":RED,color:disabled?"#aaa":"#fff",fontWeight:800,fontSize:14,cursor:disabled?"not-allowed":"pointer"}}>
      🗑 Dar de baja este producto
    </button>
  );
}

function Stock({products,onUpd,onDel,onAdjust,isAdmin,addLog,stockLog,setStockLog}) {
  const [search,setSearch]=useState("");
  const [cat,setCat]=useState("todos");
  const [editing,setEditing]=useState(null);
  const [stockTab,setStockTab]=useState("lista");
  const [adjusting,setAdjusting]=useState(null); // {pid, name, stock}
  const [adjQty,setAdjQty]=useState(0);
  const [adjReason,setAdjReason]=useState("");
  const CATS=useMemo(()=>["todos",...new Set(products.map(p=>p.category))].sort(),[products]);
  const q=search.toLowerCase();
  const shown=useMemo(()=>products.filter(p=>{
    if(cat!=="todos"&&p.category!==cat) return false;
    if(q) return p.name.toLowerCase().includes(q)||p.id.includes(q);
    return p.stock>0;
  }).slice(0,150),[products,cat,q]);
  const low=products.filter(p=>p.stock>0&&p.stock<=5);
  return (
    <div>
      {/* Sub-tabs */}
      {isAdmin && (
        <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:14,display:"flex",gap:4,boxShadow:"0 1px 4px #0001"}}>
          {[{k:"lista",l:"📋 Lista de Stock"},{k:"ajuste",l:"⚖️ Ajuste / Baja"},{k:"log",l:"📜 Movimientos"}].map(t=>(
            <button key={t.k} onClick={()=>setStockTab(t.k)} style={{flex:1,padding:"9px 14px",borderRadius:9,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:stockTab===t.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:stockTab===t.k?"#fff":"#555"}}>
              {t.l}
            </button>
          ))}
        </div>
      )}

      {isAdmin && stockTab==="ajuste" && <StockAdjust products={products} onDel={onDel} onAdjust={onAdjust} addLog={addLog}/>}
      {isAdmin && stockTab==="log" && <StockLog log={stockLog} onClear={()=>setStockLog([])}/>}
      {stockTab==="lista" && <>
      {low.length>0&&<StockAlert low={low}/>}
      <div style={{background:"#fff",borderRadius:12,padding:14,marginBottom:14,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,flex:1}}>📦 Stock ({products.filter(p=>p.stock>0).length} productos con stock)</div>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar..."
          style={{padding:"7px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:12,outline:"none",width:180}}/>
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {CATS.map(c=><button key={c} onClick={()=>setCat(c)} style={{padding:"4px 10px",borderRadius:20,border:"1.5px solid",cursor:"pointer",fontSize:11,fontWeight:600,borderColor:cat===c?RED:"#e5e5e5",background:cat===c?"#fdecea":"#fff",color:cat===c?RED:"#666"}}>{c==="todos"?"Todos":c}</button>)}
        </div>
      </div>
      {!q&&cat==="todos"&&<div style={{background:"#d6eaf8",borderRadius:8,padding:"8px 14px",fontSize:12,color:"#1a5276",marginBottom:12}}>Mostrando productos con stock. Buscá para ver el catálogo completo.</div>}
      <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px #0001",overflow:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr style={{background:"#f9f9f9"}}>
            {["Código","Producto","Categoría","Stock","P. Venta",...(isAdmin?["P. Costo","Margen"]:[]),""].map(h=><th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,color:"#888",fontSize:11,textTransform:"uppercase",letterSpacing:.5,whiteSpace:"nowrap"}}>{h}</th>)}
          </tr></thead>
          <tbody>
            {shown.length===0
              ? <tr><td colSpan={8} style={{textAlign:"center",padding:40,color:"#aaa"}}>Sin resultados.</td></tr>
              : shown.map(p=>{
                  const m=p.costPrice>0?((p.salePrice-p.costPrice)/p.costPrice*100).toFixed(0):"—";
                  return <tr key={p.id} style={{borderTop:"1px solid #f5f5f5"}}>
                    <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.id}</td>
                    <td style={{padding:"9px 12px",fontWeight:600,color:"#1a1a1a",maxWidth:260}}>{p.name}</td>
                    <td style={{padding:"9px 12px",color:"#aaa",fontSize:11}}>{p.category}</td>
                    <td style={{padding:"9px 12px"}}><SPill n={p.stock}/></td>
                    <td style={{padding:"9px 12px",fontWeight:700,color:RED}}>{fARS(p.salePrice)}</td>
                    {isAdmin&&<td style={{padding:"9px 12px",color:"#666"}}>{fARS(p.costPrice)}</td>}
                    {isAdmin&&<td style={{padding:"9px 12px",fontWeight:700,color:+m>=40?"#1e8449":"#e67e22"}}>{m}%</td>}
                    <td style={{padding:"9px 12px"}}>{isAdmin&&<button onClick={()=>setEditing(p)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11,fontWeight:600}}>✏️</button>}</td>
                  </tr>;
                })
            }
          </tbody>
        </table>
      </div>
      {editing&&<EditModal p={editing} onSave={p=>{onUpd(p);setEditing(null);}} onClose={()=>setEditing(null)}/>}
      </>}
    </div>
  );
}

function EditModal({p,onSave,onClose}) {
  const [cost,setCost]=useState(p.costPrice);
  const [sale,setSale]=useState(p.salePrice);
  const [stock,setStock]=useState(p.stock);
  const m=cost>0?((sale-cost)/cost*100).toFixed(1):"—";
  return (
    <div style={{position:"fixed",inset:0,background:"#0007",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:16}}>
      <div style={{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:440,boxShadow:"0 20px 60px #0003"}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
          <div style={{fontWeight:800,fontSize:14,color:"#1a1a1a",flex:1,marginRight:8,lineHeight:1.3}}>{p.name}</div>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:"#aaa"}}>×</button>
        </div>
        {[["Precio Costo ($)",cost,setCost],["Precio Venta ($)",sale,setSale],["Stock",stock,setStock]].map(([lbl,val,set])=>(
          <div key={lbl} style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:"#666",display:"block",marginBottom:4}}>{lbl}</label>
            <input type="number" value={val} onChange={e=>set(+e.target.value||0)} style={{...inputStyle}}/>
          </div>
        ))}
        <div style={{background:"#f0fdf4",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#1e8449",marginBottom:14}}>
          Margen: <strong>{m}%</strong> · Ganancia/u: <strong>{fARS(sale-cost)}</strong>
        </div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{padding:"8px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666"}}>Cancelar</button>
          <button onClick={()=>onSave({...p,costPrice:cost,salePrice:sale,stock})} style={{padding:"8px 14px",borderRadius:8,border:"none",background:RED,color:"#fff",cursor:"pointer",fontWeight:700}}>Guardar</button>
        </div>
      </div>
    </div>
  );
}

// ─── ALTA DE MERCADERÍA ───────────────────────────────────────────────────────
function Compras({products,onStock}) {
  const [search,setSearch]=useState("");
  const [items,setItems]=useState([]);
  const [ok,setOk]=useState(false);
  const found=useMemo(()=>{
    const q=search.toLowerCase();
    return q?products.filter(p=>p.name.toLowerCase().includes(q)||p.id.includes(q)).slice(0,50):[];
  },[products,search]);
  const addI=p=>{if(!items.find(i=>i.pid===p.id))setItems(x=>[...x,{pid:p.id,name:p.name,qty:1,cost:p.costPrice}]);};
  const remI=pid=>setItems(x=>x.filter(i=>i.pid!==pid));
  const updI=(pid,f,v)=>setItems(x=>x.map(i=>i.pid===pid?{...i,[f]:v}:i));
  const totalCost=items.reduce((s,i)=>s+i.qty*i.cost,0);
  const submit=()=>{
    items.forEach(i=>onStock(i.pid,+i.qty,+i.cost));
    setItems([]); setOk(true); setTimeout(()=>{setOk(false);},2000);
  };
  if(ok) return <div style={{textAlign:"center",padding:80}}><div style={{fontSize:60}}>📦</div><div style={{fontWeight:800,color:"#1e8449",fontSize:20,marginTop:12}}>¡Stock actualizado!</div></div>;
  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 340px",gap:18,alignItems:"start"}}>
      <div>
        <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:12,boxShadow:"0 1px 4px #0001"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>📥 Ingresar al Stock</div>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscá los productos que recibiste..."
            style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1.5px solid #e5e5e5",fontSize:13,outline:"none",boxSizing:"border-box"}}/>
          {search&&<div style={{fontSize:11,color:"#aaa",marginTop:6}}>{found.length} resultados</div>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:10}}>
          {found.length>0
            ? found.map(p=>{
                const inL=items.find(i=>i.pid===p.id);
                return <div key={p.id} style={{background:"#fff",borderRadius:10,padding:14,border:inL?"2px solid #1e8449":"2px solid transparent",boxShadow:"0 1px 4px #0001"}}>
                  <div style={{fontWeight:700,fontSize:12,color:"#1a1a1a",marginBottom:3,lineHeight:1.3}}>{p.name}</div>
                  <div style={{fontSize:10,color:"#aaa",marginBottom:8}}>{p.id} · Stock actual: <strong>{p.stock}</strong></div>
                  <button onClick={()=>addI(p)} disabled={!!inL} style={{width:"100%",padding:"7px",borderRadius:7,border:"none",fontSize:12,fontWeight:700,background:inL?"#d5f5e3":"#1e8449",color:inL?"#1a5276":"#fff",cursor:inL?"not-allowed":"pointer"}}>{inL?"✓ Agregado":"+ Agregar a la compra"}</button>
                </div>;
              })
            : <div style={{padding:20,color:"#aaa",fontSize:13}}>{search?"Sin resultados.":"Escribí el nombre del producto a ingresar."}</div>
          }
        </div>
      </div>
      <div style={{position:"sticky",top:16}}>
        <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 2px 12px #0002"}}>
          <div style={{fontWeight:800,fontSize:15,marginBottom:14}}>🧾 Detalle de Compra</div>
          {items.length===0
            ? <div style={{textAlign:"center",color:"#aaa",fontSize:12,padding:"16px 0"}}>Seleccioná productos</div>
            : items.map(it=>(
                <div key={it.pid} style={{background:"#f9f9f9",borderRadius:8,padding:12,marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                    <span style={{fontWeight:600,fontSize:12,flex:1,marginRight:6,lineHeight:1.3}}>{it.name}</span>
                    <button onClick={()=>remI(it.pid)} style={{background:"none",border:"none",cursor:"pointer",color:RED,fontSize:18,lineHeight:1}}>×</button>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <div style={{flex:1}}><div style={{fontSize:10,color:"#aaa",marginBottom:3}}>Cantidad</div>
                      <input type="number" min={1} value={it.qty} onChange={e=>updI(it.pid,"qty",+e.target.value)} style={{...inputStyle,padding:"5px 7px",fontSize:12}}/></div>
                    <div style={{flex:1}}><div style={{fontSize:10,color:"#aaa",marginBottom:3}}>P. Costo ($)</div>
                      <input type="number" value={it.cost} onChange={e=>updI(it.pid,"cost",+e.target.value)} style={{...inputStyle,padding:"5px 7px",fontSize:12}}/></div>
                  </div>
                  <div style={{fontSize:11,color:"#666",marginTop:6}}>Subtotal: <strong>{fARS(it.qty*it.cost)}</strong> · Nuevo venta: <strong style={{color:RED}}>{fARS(it.cost*1.5)}</strong></div>
                </div>
              ))
          }
          <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:16,color:RED,padding:"10px 0",borderTop:"2px solid #f5f5f5",margin:"8px 0 14px"}}>
            <span>Total compra</span><span>{fARS(totalCost)}</span>
          </div>
          <button onClick={submit} disabled={!items.length} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,background:!items.length?"#e5e5e5":"linear-gradient(135deg,#1a5e20,#1e8449)",color:!items.length?"#aaa":"#fff"}}>📦 Ingresar al Stock</button>
        </div>
      </div>
    </div>
  );
}

// ─── ADMIN PANEL ──────────────────────────────────────────────────────────────
function AdminPanel({users,setUsers,vendors,setVendors,products,setProducts,stockLog,setStockLog,notifs,setNotifs}) {
  const [section, setSection] = useState("vendors");

  const SECTIONS = [
    {k:"vendors", label:"Vendedores",       icon:"👥"},
    {k:"users",   label:"Usuarios",         icon:"🔐"},
    {k:"excel",   label:"Lista de Precios", icon:"📊"},
    {k:"notifcfg",label:"Notificaciones",   icon:"🔔"},
  ];

  return (
    <div>
      <div style={{background:"#fff",borderRadius:12,padding:4,marginBottom:16,display:"flex",gap:4,boxShadow:"0 1px 4px #0001",flexWrap:"wrap"}}>
        {SECTIONS.map(s=>(
          <button key={s.k} onClick={()=>setSection(s.k)} style={{flex:1,minWidth:120,padding:"10px 16px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,background:section===s.k?`linear-gradient(135deg,${REDD},${RED})`:"transparent",color:section===s.k?"#fff":"#555",transition:"all .15s"}}>
            {s.icon} {s.label}
          </button>
        ))}
      </div>
      {section==="vendors"  && <VendorsPanel vendors={vendors} setVendors={setVendors}/>}
      {section==="users"    && <UsersPanel   users={users}     setUsers={setUsers}/>}
      {section==="excel"    && <ExcelPanel   products={products} setProducts={setProducts}/>}
      {section==="notifcfg" && <NotifConfig  users={users} setUsers={setUsers} notifs={notifs} setNotifs={setNotifs}/>}
    </div>
  );
}

// ── Vendors ───────────────────────────────────────────────────────────────────
function VendorsPanel({vendors,setVendors}) {
  const [newName, setNewName] = useState("");
  const [editing, setEditing] = useState(null);
  const [editVal, setEditVal] = useState("");

  const add = () => {
    const n = newName.trim();
    if(!n) return;
    if(vendors.includes(n)){alert("Ya existe ese vendedor");return;}
    setVendors(v=>[...v,n]);
    setNewName("");
  };
  const [confirmDel, setConfirmDel] = useState(null);
  const remove = (v) => setConfirmDel(v);
  const doRemove = () => { setVendors(vs=>vs.filter(x=>x!==confirmDel)); setConfirmDel(null); };
  const saveEdit = (old) => {
    const n = editVal.trim();
    if(!n) return;
    setVendors(vs=>vs.map(x=>x===old?n:x));
    setEditing(null);
  };

  return (
    <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001",maxWidth:520}}>
      <div style={{fontWeight:800,fontSize:16,marginBottom:16}}>👥 Gestión de Vendedores</div>
      <div style={{display:"flex",gap:8,marginBottom:20}}>
        <input value={newName} onChange={e=>setNewName(e.target.value)}
          onKeyDown={e=>e.key==="Enter"&&add()}
          placeholder="Nombre del nuevo vendedor"
          style={{...inputStyle,flex:1}}/>
        <button onClick={add} style={{padding:"8px 18px",borderRadius:8,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:13,whiteSpace:"nowrap"}}>+ Agregar</button>
      </div>
      {vendors.length===0
        ? <div style={{textAlign:"center",padding:30,color:"#aaa"}}>No hay vendedores cargados.</div>
        : vendors.map(v=>(
          <div key={v} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,border:"1.5px solid #f0f0f0",marginBottom:8,background:"#fafafa"}}>
            <span style={{fontSize:20}}>👤</span>
            {editing===v
              ? <>
                  <input value={editVal} onChange={e=>setEditVal(e.target.value)}
                    onKeyDown={e=>{if(e.key==="Enter")saveEdit(v);if(e.key==="Escape")setEditing(null);}}
                    style={{...inputStyle,flex:1,fontSize:13}} autoFocus/>
                  <button onClick={()=>saveEdit(v)} style={{padding:"5px 12px",borderRadius:7,border:"none",background:"#1e8449",color:"#fff",fontWeight:700,cursor:"pointer",fontSize:12}}>✓</button>
                  <button onClick={()=>setEditing(null)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>✕</button>
                </>
              : <>
                  <span style={{flex:1,fontWeight:600,fontSize:14}}>{v}</span>
                  <button onClick={()=>{setEditing(v);setEditVal(v);}} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:12}}>✏️ Editar</button>
                  {confirmDel===v
                    ? <div style={{display:"flex",alignItems:"center",gap:5,background:"#fdecea",borderRadius:8,padding:"4px 8px",border:`1px solid ${RED}44`}}>
                        <span style={{fontSize:11,color:RED,fontWeight:600,whiteSpace:"nowrap"}}>¿Eliminar?</span>
                        <button onClick={doRemove} style={{padding:"3px 9px",borderRadius:6,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:11}}>Sí</button>
                        <button onClick={()=>setConfirmDel(null)} style={{padding:"3px 8px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11}}>No</button>
                      </div>
                    : <button onClick={()=>remove(v)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:12}}>🗑</button>
                  }
                </>
            }
          </div>
        ))
      }
    </div>
  );
}

// ── Users ─────────────────────────────────────────────────────────────────────
function UsersPanel({users,setUsers}) {
  const [form, setForm] = useState({username:"",password:"",name:"",role:"vendedor"});
  const [editing, setEditing] = useState(null);
  const [showPass, setShowPass] = useState({});

  const startEdit = (u) => { setEditing(u.id); setForm({username:u.username,password:u.password,name:u.name,role:u.role}); };
  const cancelEdit = () => { setEditing(null); setForm({username:"",password:"",name:"",role:"vendedor"}); };

  const save = () => {
    if(!form.username.trim()||!form.password.trim()||!form.name.trim()){alert("Completá todos los campos");return;}
    if(editing) {
      setUsers(us=>us.map(u=>u.id===editing?{...u,...form}:u));
    } else {
      if(users.find(u=>u.username===form.username.trim())){alert("Ese usuario ya existe");return;}
      setUsers(us=>[...us,{id:genId(),...form,username:form.username.trim(),name:form.name.trim()}]);
    }
    cancelEdit();
  };
  const remove = (id) => {
    if(users.filter(u=>u.role==="admin").length===1&&users.find(u=>u.id===id)?.role==="admin"){alert("Debe haber al menos un administrador");return;}
    if(confirm("¿Eliminar este usuario?")) setUsers(us=>us.filter(u=>u.id!==id));
  };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      {/* Form */}
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>{editing?"✏️ Editar usuario":"➕ Nuevo usuario"}</div>
        <Field label="Nombre completo">
          <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Ej: María García" style={inputStyle}/>
        </Field>
        <Field label="Usuario (para login)">
          <input value={form.username} onChange={e=>setForm(f=>({...f,username:e.target.value}))} placeholder="Ej: maria" style={inputStyle}/>
        </Field>
        <Field label="Contraseña">
          <div style={{position:"relative"}}>
            <input type={showPass.form?"text":"password"} value={form.password}
              onChange={e=>setForm(f=>({...f,password:e.target.value}))}
              placeholder="Contraseña segura" style={{...inputStyle,paddingRight:40}}/>
            <button onClick={()=>setShowPass(s=>({...s,form:!s.form}))} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:15,color:"#aaa"}}>{showPass.form?"🙈":"👁"}</button>
          </div>
        </Field>
        <Field label="Email (para notificaciones)">
          <input type="email" value={form.email||""} onChange={e=>setForm(f=>({...f,email:e.target.value}))}
            placeholder="correo@ejemplo.com" style={inputStyle}/>
        </Field>
        <Field label="Rol">
          <select value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))} style={{...inputStyle,cursor:"pointer"}}>
            <option value="vendedor">Vendedor</option>
            <option value="admin">Administrador</option>
          </select>
        </Field>
        <div style={{display:"flex",gap:8,marginTop:4}}>
          <button onClick={save} style={{flex:1,padding:"9px",borderRadius:8,border:"none",background:RED,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:13}}>{editing?"Guardar cambios":"Crear usuario"}</button>
          {editing&&<button onClick={cancelEdit} style={{padding:"9px 14px",borderRadius:8,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontWeight:600,color:"#666",fontSize:13}}>Cancelar</button>}
        </div>
      </div>
      {/* List */}
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:16}}>🔐 Usuarios ({users.length})</div>
        {users.map(u=>(
          <div key={u.id} style={{padding:"12px 14px",borderRadius:10,border:`1.5px solid ${editing===u.id?"#c0392b":"#f0f0f0"}`,marginBottom:8,background:editing===u.id?"#fdecea":"#fafafa"}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:22}}>{u.role==="admin"?"👑":"👤"}</span>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:13}}>{u.name}</div>
                <div style={{fontSize:11,color:"#888"}}>@{u.username} · <span style={{color:u.role==="admin"?RED:"#1a5276",fontWeight:600}}>{u.role==="admin"?"Admin":"Vendedor"}</span>{u.email&&<span style={{color:"#aaa",marginLeft:6}}>· {u.email}</span>}</div>
              </div>
              <button onClick={()=>startEdit(u)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #e5e5e5",background:"#fff",cursor:"pointer",fontSize:11}}>✏️</button>
              <button onClick={()=>remove(u.id)} style={{padding:"4px 10px",borderRadius:6,border:"1.5px solid #fcc",background:"#fff",color:RED,cursor:"pointer",fontSize:11}}>🗑</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Excel Import ──────────────────────────────────────────────────────────────
function ExcelPanel({products,setProducts}) {
  const fileRef = useRef();
  const [preview, setPreview] = useState(null);
  const [status, setStatus] = useState(null);
  const [mode, setMode] = useState("update");

  const parseExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        // raw:true para obtener valores numéricos reales, no strings formateados
        // cellFormula:false para que SheetJS use el valor cacheado (.v) y no el texto de la fórmula
        const wb = XLSX.read(e.target.result, {
          type:"array",
          cellFormula: false,
          cellNF: false,
          cellHTML: false,
        });
        const ws = wb.Sheets[wb.SheetNames[0]];

        // Leer rango de la hoja
        const range = XLSX.utils.decode_range(ws["!ref"] || "A1:G1000");
        const totalRows = range.e.r; // última fila con datos

        if(totalRows < 1){setStatus({type:"error",msg:"El archivo parece estar vacío."});return;}

        // Función que extrae el valor numérico REAL de una celda (usa .v = valor cacheado)
        // Esto es clave para fórmulas: .v tiene el resultado que Excel calculó y guardó
        const cellNum = (r, c) => {
          const addr = XLSX.utils.encode_cell({r, c});
          const cell = ws[addr];
          if(!cell || cell.v === undefined || cell.v === null || cell.v === "") return null;
          const n = typeof cell.v === "number" ? cell.v
            : parseFloat(String(cell.v).replace(/[$\s]/g,"").replace(/\./g,"").replace(",","."));
          return isNaN(n) || n === 0 ? null : n;
        };
        const cellStr = (r, c) => {
          const addr = XLSX.utils.encode_cell({r, c});
          const cell = ws[addr];
          if(!cell || cell.v === undefined) return "";
          return String(cell.v).trim();
        };

        // Detectar fila de encabezado (primera fila con "CODIGO" o "CÓDIGO" en alguna celda)
        let headerRow = 0;
        for(let r = 0; r <= Math.min(5, totalRows); r++) {
          for(let c = 0; c <= 10; c++) {
            const v = cellStr(r, c).toUpperCase()
              .normalize("NFD").replace(/[\u0300-\u036f]/g,"");
            if(v.includes("CODIGO") || v.includes("COD")) { headerRow = r; break; }
          }
        }

        // Detectar columnas por nombre del encabezado (con normalización de tildes)
        const norm = s => String(s).toUpperCase().trim()
          .normalize("NFD").replace(/[\u0300-\u036f]/g,"");

        const COL = { codigo:-1, descripcion:-1, precioIVA:-1, precioOferta:-1, precioFinal:-1, fecha:-1 };
        for(let c = range.s.c; c <= range.e.c; c++) {
          const h = norm(cellStr(headerRow, c));
          if(h.includes("CODIGO") || h === "COD" || h === "ID") COL.codigo = c;
          else if(h.includes("DESCRIP") || h.includes("NOMBRE"))       COL.descripcion = c;
          else if(h.includes("CON IVA") || (h.includes("IVA") && !h.includes("OFERTA") && !h.includes("FINAL"))) COL.precioIVA = c;
          else if(h.includes("OFERTA"))  COL.precioOferta = c;
          else if(h.includes("FINAL"))   COL.precioFinal  = c;
          else if(h.includes("FECHA") || h.includes("ULTIMA") || h.includes("ACT")) COL.fecha = c;
        }

        // Fallback posición fija si alguna columna no se detectó por nombre
        // Estructura conocida: A(0)=Código B(1)=Descripción C(2)=P.IVA D(3)=P.Oferta E(4)=Fecha F(5)=Código G(6)=P.Final
        if(COL.codigo      < 0) COL.codigo      = 0;
        if(COL.descripcion < 0) COL.descripcion = 1;
        if(COL.precioIVA   < 0) COL.precioIVA   = 2;
        if(COL.precioOferta< 0) COL.precioOferta = 3;
        if(COL.fecha       < 0) COL.fecha        = 4;
        if(COL.precioFinal < 0) COL.precioFinal  = 6;

        // Diagnóstico de columnas detectadas
        const colLetter = c => c < 0 ? "—" : String.fromCharCode(65 + c);
        const detectedCols = {
          "Código":        colLetter(COL.codigo),
          "Descripción":   colLetter(COL.descripcion),
          "Precio c/IVA":  colLetter(COL.precioIVA),
          "Precio Oferta": colLetter(COL.precioOferta),
          "Fecha":         colLetter(COL.fecha),
          "Precio Final":  colLetter(COL.precioFinal),
        };

        // Parsear filas de datos (después del encabezado)
        const parsed = [];
        for(let r = headerRow + 1; r <= totalRows; r++) {
          const id = cellStr(r, COL.codigo);
          if(!id) continue;

          // Leer precios directamente desde objeto de celda (.v = valor real/calculado)
          const pIVA    = cellNum(r, COL.precioIVA);    // col C
          const pOferta = cellNum(r, COL.precioOferta); // col D
          const pFinal  = cellNum(r, COL.precioFinal);  // col G (fórmula)

          parsed.push({
            id,
            name:         cellStr(r, COL.descripcion),
            precioIVA:    pIVA,
            precioOferta: pOferta,
            precioFinal:  pFinal,
            fecha:        cellStr(r, COL.fecha),
          });
        }

        if(parsed.length === 0){setStatus({type:"error",msg:"No se encontraron datos. Verificá que el archivo tenga encabezados en la primera fila."});return;}

        setPreview({rows:parsed.slice(0,10), total:parsed.length, all:parsed, cols:COL, detectedCols});
        setStatus(null);
      } catch(err) {
        setStatus({type:"error",msg:"Error al leer el archivo: "+err.message});
        setPreview(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const applyImport = () => {
    if(!preview) return;
    const updated=[], notFound=[];
    const newProds = products.map(p=>({...p}));

    // Regla de precio: si Oferta (D) tiene valor numérico mayor a 0 → se usa D, sino → se usa C (IVA)
    // Aplica igual para precio de venta Y precio de costo
    const resolvePrice = (row) => {
      const oferta = row.precioOferta;
      const iva    = row.precioIVA;
      // D tiene prioridad si es un número válido > 0
      if(oferta !== null && oferta !== undefined && !isNaN(oferta) && oferta > 0) return oferta;
      return iva;
    };

    preview.all.forEach(row=>{
      const idx = newProds.findIndex(p=>p.id===row.id);
      const precio = resolvePrice(row); // D si existe, sino C
      if(idx>=0){
        // Precio de venta: col G (precioFinal, fórmula) tiene prioridad;
        // si no viene, se aplica la misma regla D > C
        if(row.precioFinal!==null) newProds[idx].salePrice = row.precioFinal;
        else if(precio!==null)     newProds[idx].salePrice = precio;
        // Precio de costo: misma regla D > C
        if(precio!==null)          newProds[idx].costPrice = precio;
        if(row.name) newProds[idx].name = row.name;
        updated.push(row.id);
      } else {
        notFound.push(row.id);
      }
    });

    if(mode==="full") {
      // also add new products not in catalog
      preview.all.forEach(row=>{
        if(!newProds.find(p=>p.id===row.id)) {
          const precio = resolvePrice(row);
          newProds.push({
            id:row.id, name:row.name||row.id,
            costPrice:precio||0,
            salePrice:row.precioFinal||precio||0,
            category:"Importado", stock:0
          });
        }
      });
    }

    setProducts(newProds);
    setStatus({type:"success",msg:`✅ ${updated.length} productos actualizados.${notFound.length>0?` ${notFound.length} códigos no encontrados en el catálogo.`:""}`});
    setPreview(null);
    if(fileRef.current) fileRef.current.value="";
  };

  return (
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,alignItems:"start"}}>
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:16,marginBottom:6}}>📊 Importar Lista de Precios</div>
        <div style={{fontSize:13,color:"#666",marginBottom:16,lineHeight:1.6}}>
          El sistema lee automáticamente las columnas:<br/>
          <strong>CÓDIGO · DESCRIPCIÓN · PRECIO CON IVA · PRECIO OFERTA · FECHA ULTIMA ACTUALIZACIÓN · PRECIO FINAL</strong>
        </div>

        <div style={{marginBottom:14}}>
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:6}}>Modo de importación</div>
          <div style={{display:"flex",gap:8}}>
            {[{v:"update",l:"Actualizar precios",desc:"Solo actualiza productos existentes por código"},{v:"full",l:"Importación completa",desc:"Actualiza existentes y agrega nuevos"}].map(opt=>(
              <div key={opt.v} onClick={()=>setMode(opt.v)} style={{flex:1,padding:"10px 12px",borderRadius:10,border:`2px solid ${mode===opt.v?RED:"#e5e5e5"}`,background:mode===opt.v?"#fdecea":"#fafafa",cursor:"pointer"}}>
                <div style={{fontWeight:700,fontSize:12,color:mode===opt.v?RED:"#555"}}>{opt.l}</div>
                <div style={{fontSize:10,color:"#aaa",marginTop:2}}>{opt.desc}</div>
              </div>
            ))}
          </div>
        </div>

        <div onClick={()=>fileRef.current?.click()} style={{border:"2px dashed #e5e5e5",borderRadius:12,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:"#fafafa",marginBottom:14,transition:"border-color .2s"}}
          onMouseEnter={e=>e.currentTarget.style.borderColor=RED}
          onMouseLeave={e=>e.currentTarget.style.borderColor="#e5e5e5"}>
          <div style={{fontSize:36,marginBottom:8}}>📂</div>
          <div style={{fontWeight:700,color:"#555",fontSize:14}}>Hacé clic o arrastrá tu archivo Excel</div>
          <div style={{fontSize:12,color:"#aaa",marginTop:4}}>.xlsx · .xls · .csv</div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}}
            onChange={e=>{if(e.target.files[0])parseExcel(e.target.files[0]);}}/>
        </div>

        {status&&<div style={{padding:"10px 14px",borderRadius:8,background:status.type==="error"?"#fdecea":"#d5f5e3",color:status.type==="error"?RED:"#1e8449",fontSize:13,fontWeight:600,marginBottom:12}}>{status.msg}</div>}

        {preview&&(
          <button onClick={applyImport} style={{width:"100%",padding:"11px",borderRadius:10,border:"none",background:`linear-gradient(135deg,${REDD},${RED})`,color:"#fff",fontWeight:800,fontSize:14,cursor:"pointer"}}>
            📥 Aplicar importación ({preview.total} productos)
          </button>
        )}
      </div>

      {/* Preview */}
      <div style={{background:"#fff",borderRadius:12,padding:24,boxShadow:"0 1px 4px #0001"}}>
        <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>
          {preview ? `👁 Vista previa (${preview.total} filas)` : "📋 Instrucciones"}
        </div>
        {preview?.detectedCols && (
          <div style={{background:"#f0fdf4",border:"1px solid #a7f3d0",borderRadius:8,padding:"10px 12px",marginBottom:12}}>
            <div style={{fontWeight:700,fontSize:11,color:"#065f46",marginBottom:6}}>✅ Columnas detectadas:</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {Object.entries(preview.detectedCols).map(([k,v])=>(
                <span key={k} style={{background:"#fff",border:"1px solid #d1fae5",borderRadius:6,padding:"3px 8px",fontSize:11,color:"#065f46"}}>
                  <strong>{v}</strong> → {k}
                </span>
              ))}
            </div>
          </div>
        )}
        {!preview
          ? <div style={{fontSize:13,color:"#666",lineHeight:1.8}}>
              <div style={{background:"#f9f9f9",borderRadius:8,padding:"12px 16px",fontFamily:"monospace",fontSize:12,lineHeight:2}}>
                <div><strong>Col A:</strong> CÓDIGO</div>
                <div><strong>Col B:</strong> DESCRIPCIÓN</div>
                <div><strong>Col C:</strong> PRECIO CON IVA</div>
                <div><strong>Col D:</strong> PRECIO OFERTA</div>
                <div><strong>Col E:</strong> FECHA ULTIMA ACTUALIZACIÓN</div>
                <div><strong>Col F:</strong> PRECIO FINAL</div>
              </div>
              <div style={{marginTop:12,fontSize:12,color:"#aaa"}}>El sistema detecta automáticamente las columnas por nombre. Las columnas no tienen que estar en un orden específico.</div>
            </div>
          : <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{background:"#f9f9f9"}}>
                  {["Código","Descripción","P. IVA","P. Oferta","P. Final","Fecha"].map(h=><th key={h} style={{padding:"7px 8px",textAlign:"left",fontWeight:700,color:"#888",whiteSpace:"nowrap"}}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {preview.rows.map((r,i)=>(
                    <tr key={i} style={{borderTop:"1px solid #f5f5f5"}}>
                      <td style={{padding:"6px 8px",fontWeight:600,color:"#444"}}>{r.id}</td>
                      <td style={{padding:"6px 8px",color:"#555",maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</td>
                      <td style={{padding:"6px 8px",color:"#666"}}>{r.precioIVA!=null?fARS(r.precioIVA):"—"}</td>
                      <td style={{padding:"6px 8px",color:"#666"}}>{r.precioOferta!=null?fARS(r.precioOferta):"—"}</td>
                      <td style={{padding:"6px 8px",fontWeight:700,color:RED}}>{r.precioFinal!=null?fARS(r.precioFinal):"—"}</td>
                      <td style={{padding:"6px 8px",color:"#aaa",fontSize:10}}>{r.fecha}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {preview.total>10&&<div style={{fontSize:11,color:"#aaa",marginTop:8,textAlign:"center"}}>... y {preview.total-10} filas más</div>}
            </div>
        }
      </div>
    </div>
  );
}
