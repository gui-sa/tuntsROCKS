"strict mode"
const fetch = require('node-fetch');

let url = "https://restcountries.com/v3.1/all";

let settings = { method: "Get" };

fetch(url, settings)
    .then( res => res.json())
    .then((json) => {
        const xl = require('excel4node');
        var wb = new xl.Workbook();
        var ws = wb.addWorksheet('OUTPUT',{
            'printOptions': {
                'centerHorizontal': 1,
                'centerVertical': 1,
                'printGridLines': 0,
                'printHeadings': 0
        
            }
        });
        wb.write('OUTPUT.xlsx'); 

        let titulo1_modifier ={
            bold: true,
            color: '#4F4F4F',
            size: 16,
            name: 'Arial',};

        let titulo1_style = wb.createStyle({
            alignment: {
                horizontal: 'center',
                vertical: 'center',},
                border: { 
                    left: {
                        style: 'thin',
                        color: '000000' 
                    },
                    right: {
                        style: 'thin',
                        color: '000000' 
                    },
                    top: {
                        style: 'thin',
                        color: '000000' 
                    },
                    bottom: {
                        style: 'thin',
                        color: '000000' 
                    }
            }});

        let titulo2_modifier ={
            bold: true,
            color: '#808080',
            size: 12,
            name: 'Arial',};
        
        let titulo2_style = wb.createStyle({
            alignment: {
                horizontal: 'left',
                vertical: 'center',},
                border: { 
                    left: {
                        style: 'thin',
                        color: '000000' 
                    },
                    right: {
                        style: 'thin',
                        color: '000000' 
                    },
                    top: {
                        style: 'thin',
                        color: '000000' 
                    },
                    bottom: {
                        style: 'thin',
                        color: '000000' 
                    }
            }});

        let text_style = wb.createStyle({
            alignment: {
                horizontal: 'left',
                vertical: 'center',},
                border: { 
                    left: {
                        style: 'thin',
                        color: '000000' 
                    },
                    right: {
                        style: 'thin',
                        color: '000000' 
                    },
                    top: {
                        style: 'thin',
                        color: '000000' 
                    },
                    bottom: {
                        style: 'thin',
                        color: '000000' 
                    }
            }});
    
        let text_modifier ={
            bold: false,
            color: '000000',
            size: 12,
            name: 'Arial',};

        let number_style = wb.createStyle({
            alignment: {
                horizontal: 'left',
                vertical: 'center',},
                border: { 
                    left: {
                        style: 'thin',
                        color: '000000' 
                    },
                    right: {
                        style: 'thin',
                        color: '000000' 
                    },
                    top: {
                        style: 'thin',
                        color: '000000' 
                    },
                    bottom: {
                        style: 'thin',
                        color: '000000' 
                    }
            },
            font: {
                bold: false,
                color: '000000',
                size: 12,
                name: 'Arial',
            },
            numberFormat: '#.##0,00' });

        ws.cell(1, 1, 1 , 4, true).string([titulo1_modifier,'Countries List']).style(titulo1_style);  
        ws.cell(2, 1).string([titulo2_modifier,'Name']).style(titulo2_style);  
        ws.cell(2, 2).string([titulo2_modifier,'Capital']).style(titulo2_style);  
        ws.cell(2, 3).string([titulo2_modifier,'Area']).style(titulo2_style);  
        ws.cell(2, 4).string([titulo2_modifier,'Currencies']).style(titulo2_style);  

        let width_name = 0;
        let width_capital = 0;
        let width_area = 0;
        let width_currencies = 0;
        
        for (let i=0; i< json.length;i++){
            //Name
            ws.cell(3+i, 1).string([text_modifier, json[i].name.official]).style(text_style);  
            
            //Capital
            if (typeof (json[i].capital) == 'undefined') {
                ws.cell(3+i, 2).string([text_modifier, '-']).style(text_style);  
            }else{
                ws.cell(3+i, 2).string([text_modifier, json[i].capital[0]]).style(text_style);  
            }
            
            //Area
            ws.cell(3+i, 3).number(json[i].area).style(number_style);  

            //Currencies
            if (typeof (json[i].currencies) == 'undefined') {
                ws.cell(3+i, 4).string([text_modifier, '-']).style(text_style);  
            }else{
                var currencies_text = '';
                for(let j=0 ;  j < Object.keys(json[i].currencies).length;j++){
                    currencies_text += Object.keys(json[i].currencies)[j] + ', ';
                }
                currencies_text = currencies_text.substring(0,currencies_text.length-2);
                ws.cell(3+i, 4).string([text_modifier, currencies_text]).style(text_style);  
            }
            ws.column(1).setWidth(100);
        }
        console.log("File Successfully created")
});