import XLSX from 'sheetjs-style';    

const wb = XLSX.utils.book_new();

const file = XLSX.readFile('orden2.xlsx');
const sheets = file.SheetNames
const ws = file.Sheets[sheets[0]];
const wsTable = {};

const data = {
    proceso: "INAGUJA-UC-CD-2022-0052",
    fecha: "Fecha de emision: 12/12/2022",
    tipo: "ORDEN DE COMPRA",
    numero: "INAGUJA-2022-0050",
    descripcion: "CONTRATACION DE SERVICIOS LEGALES PARA PROCESO DE LOCALES DE LA INSTITUCION, DESTINADO A MIPYME",
    modalidad: "Compra por debajo del umbral",
    datos_proveedor: {
        razon: "POLITICOS CORRUPTOS",
        rnc: "000-00000-0",
        nombre: "POLITICOS CORRUPTOS",
        domicilio: "CALLE CORRUPCION ESQUINA ATRACO",   
        telefono: "849-123-4567" 
    },
    datos_contrato:{
        anticipo: 0,
        forma_de_pago: "Transferencia",
        plazo: "15 dias",
        monto: 500000.00,
        moneda: "DOP"
    }
}

const cellsMiddleAlign = ['G3','G4','G5','C8','C9','C11','C44','A38','A42','F38','F42'];
const cellsLeftAlign = ['C13','C14','C15','C19','C20','C21','C22','C23','C27','C28','C29','C30','C31'];
const cellsBlueLeft = ['A17','A25','A48'];
const cellsSign = ['A37','B37','C37','D37','F37','G37','H37','I37','A41','B41','C41','D41','F41','G41','H41','I41']
const cellsDetailTable = ['A50','B50','C50','D50','E50','F50','G50','H50','I50',];

//SIGN LINES
ws.A37 = {t: "", v: ""}, ws.B37 = {t: "", v: ""}, ws.C37 = {t: "", v: ""}, ws.D37 = {t: "", v: ""};
ws.A41 = {t: "", v: ""}, ws.B41 = {t: "", v: ""}, ws.C41 = {t: "", v: ""}, ws.D41 = {t: "", v: ""};

ws.F37 = {t: "", v: ""}, ws.G37 = {t: "", v: ""}, ws.H37 = {t: "", v: ""}, ws.I37 = {t: "", v: ""}; 
ws.F41 = {t: "", v: ""}, ws.G41 = {t: "", v: ""}, ws.H41 = {t: "", v: ""}, ws.I41 = {t: "", v: ""}; 

//TOP RIGHT BOX BORDER
ws.H3 = {t:"", v:"", s: {border:{top:{style:'medium'}}}}, ws.I3 = {t:"", v:"", s: {border:{top:{style:'medium'},right:{style:'medium'}}}}
ws.H4 = {t:"", v:"", s: {border:{bottom:{style:'medium'}}}}, ws.I4 = {t:"", v:"", s: {border:{right:{style:'medium'},bottom:{style:'medium'}}}}
ws.H1 = {...ws.H1, s:{alignment:{horizontal:"right"}}}

//GENERAL
ws.G4.v = data.proceso;
ws.G5.v = data.fecha;
ws.C9.v = data.tipo;
ws.C13.v = data.numero;
ws.C14.v = data.descripcion;
ws.C15.v = data.modalidad;
ws.C44.v = data.proceso;
//PROVEEDOR
ws.C19.v = data.datos_proveedor.razon;
ws.C20.v = data.datos_proveedor.rnc;
ws.C21.v = data.datos_proveedor.nombre;
ws.C22.v = data.datos_proveedor.domicilio;
ws.C23.v = data.datos_proveedor.telefono;
//CONTRATO
ws.C27.v = data.datos_contrato.anticipo;
ws.C28.v = data.datos_contrato.forma_de_pago;
ws.C29.v = data.datos_contrato.plazo;
ws.C30.v = data.datos_contrato.monto;
ws.C31.v = data.datos_contrato.moneda;


//STYLES
cellsMiddleAlign.forEach(cell => {
    if( cell === 'G3'){
        ws[cell].s ={
            font: {
                bold: true,
                color: {rgb: "FFFFFF"}
            },
            fill: {
                fgColor: {rgb: "1155CC"}
            },
            alignment: {
                vertical: 'center',
                horizontal: 'center'
            },
            border: {
                top: {style: 'medium'},
                left: {style: 'medium'},
            }
        }  
        return
    }   
    if( cell === "G4"){
        ws[cell].s = {
            font: {
                bold: true
            },
            alignment: {
                vertical: 'center',
                horizontal: 'center'
            },
            border: {
                left: {style: 'medium'},
                bottom: {style: 'medium'}
            }
        }
        return
    }
    if( cell === "G5"){
        ws[cell].s = {
            font: {
                bold: true
            },
            alignment: {
                vertical: 'center',
                horizontal: 'center'
            },
        }
        return
    }
    if( cell === 'C44'){
        ws[cell].s ={
            alignment: {
                vertical: 'center',
                horizontal: 'center'
            }
        }  
        return
    }
    ws[cell].s = {
        font: {
            bold: true
        },
        alignment: {
            vertical: 'center',
            horizontal: 'center'
        }
    }
})

cellsLeftAlign.forEach(cell => {
    ws[cell].s ={
        font: {
            bold: true
        },
        alignment: {
            vertical: 'center',
            horizontal: 'left',
            wrapText: true
        }
    }
})

cellsBlueLeft.forEach(cell => {
    ws[cell].s ={
        font: {
            bold: true,
            color: {rgb: "FFFFFF"}
        },
        fill: {
            fgColor: {rgb: "1155CC"}
        },
        alignment: {
            vertical: 'center',
            horizontal: 'left'
        }
    }
})

cellsSign.forEach(cell => {
    ws[cell].s =  {
        font: {
            bold: true
        },
        alignment: {
            vertical: 'center',
            horizontal: 'center'
        },
        border: {
            bottom: {
                style: "thick"
            }
        }
    } 
})

ws.D50 = {t:"", v:""}, ws.E50 = {t:"", v:""},

cellsDetailTable.forEach(cell => {
    ws[cell].s ={
        font: {
            bold: true,
            color: {rgb: "FFFFFF"}
        },
        fill: {
            fgColor: {rgb: "1155CC"}
        },
        alignment: {
            vertical: 'center',
            horizontal: 'center'
        },
        border: {
            top:{style:'medium'},
            bottom:{style:'medium'},
            left:{style:'medium'},
            right:{style:'medium'},
        }
    }
})

ws['A34'].s = {
    font: {
        bold: true,
        color: {rgb: "FFFFFF"}
    },
    fill: {
        fgColor: {rgb: "000000"}
    },
    alignment: {
        vertical: 'center',
        horizontal: 'left'
    }
}
ws.C11.s.font.bold = false;
ws.A14 = {...ws.A14, s: {alignment: {vertical: 'center'}}};

ws['!rows'] = []
ws['!cols'] = [ 
    {wch: 8}, //A
    {wch: 12}, //B
    {wch: 8}, //C
    , //D
    , //E
    {wch: 8}, //F 
    {wch: 6}, //G
    {wch: 11}, //H
    {wch: 9}, //I
]

ws['!margins'] = {
    left: 0.4,
    right: 0.4,
    top: 0.4,
    bottom: 0.4,
    header: 0,
    footer: 0
}
ws.H51 = {...ws.H51, z: "4"}
ws.I51 = {...ws.I51, z: "4"}
ws.C30 = {...ws.C30, z: "4"}
ws.C27 = {...ws.C27, z: "10"}

ws['!rows'][13] = {hpx: (ws.C14.v.length/45)*15}

//H51 z:"4", I51 z:"4"



XLSX.utils.book_append_sheet(wb,ws,'Orden');
XLSX.writeFile(wb,'orden2.xlsx')