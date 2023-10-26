/**
 * PleasanterExcelMaker
 *
 * @author     Satoru Sugamura <s-sugamura@senet.co.jp>
 * @copyright  2023 System Engineering and Service, Inc. All Rights Reserved.
 * @license    GNU Affero General Public License v3.0
 * @version    1.0
 * @date       2023/10/24
**/
class PleasanterExcelMaker{
    constructor(){
        this.Debug = 0;
        if(this.Debug==1)console.log("constructor()");
        this.getDefaultView = {"ApiDataType": "KeyValues", "ApiColumnKeyDisplayType": "ColumnName"}
        this.prefix = "";
    }
    ApiGet(id, data){
        return new Promise((resolve, reject) => {
            let me = this;
            if(me.Debug==1)console.log("ApiGet()");
            id == undefined || id == 0 ? id = $p.id() : id = id;
            data == undefined ? data = {"ApiVersion":1.0, "View": me.getDefaultView} : data = data;
            $p.apiGet({
                id: id,
                data: data
            }).done(function(data, textStatus, jqXHR){
                resolve({"data":data, "jqXHR":jqXHR, "textStatus":textStatus});                
            }).fail(function(jqXHR, textStatus, errorThrown){
                reject({"jqXHR":jqXHR, "textStatus":textStatus});
            });
        });
    }
    Ajax(url, data){
        return new Promise((resolve, reject) => {
            let me = this;
            if(me.Debug==1)console.log("Ajax()");
            url == undefined || url == "" ? url = me.prefix + "/api/extended/sql" : url = url;
            data == undefined ? data = {"ApiVersion":1.0, "View": this.getDefaultView} : data = data;
            $.ajax({
                type: "post",
                url: url,
                dataType: "json",
                contentType: "application/json",
                data: JSON.stringify(data)
            }).done(function(data, textStatus, jqXHR){
                resolve({"data":data, "jqXHR":jqXHR, "textStatus":textStatus});                
            }).fail(function(jqXHR, textStatus, errorThrown){
                reject({"jqXHR":jqXHR, "textStatus":textStatus});
            });
        });
    }
    ExcelDownloadFile(excelurl){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("ExcelDownloadFile()");
            excelurl == undefined ? excelurl = me.prefix + "/"+String($p.id())+".xlsx" : excelurl = excelurl;
            const Http = new XMLHttpRequest();
            Http.open("GET", excelurl+"?v="+Date.now());
            Http.responseType = 'blob';
            Http.send();
            Http.onload = function(event){
                if(me.Debug==1)console.log("Excel Download Done.");
                if(me.Debug==1)console.log("Status:"+Http.status);
                if(Http.status == 200){
                    const workbook = new ExcelJS.Workbook();
                    me.LoadBlob(Http.response, workbook).then(function(workbook){
                        resolve(workbook);
                    });    
                } else {
                    reject({status:Http.status});
                }
            }
            Http.onerror = function() {
                if(me.Debug==1)console.log("Excel Download Error.");
                reject({status:Http.status});
            }
        });
    }
    ExcelDownloadGuid(guid){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("ExcelDownloadGuid()");
            guid == undefined ? guid = "0000000000000000" : guid = guid;
            $.ajax({
                type: "post",
                url: me.prefix + "/api/binaries/"+guid+"/get",
                dataType: "json",
                contentType: "application/json"
            }).done(function(data, textStatus, jqXHR){
                if(me.Debug==1)console.log("Excel Download Done.");
                const ext = data.Response.Extension;
                if(ext==".xlsx" || ext==".xls"){
                    const blob = me.ToBlob(data.Response.Base64, data.Response.ContentType);
                    const workbook = new ExcelJS.Workbook();
                    me.LoadBlob(blob, workbook).then(function(workbook){
                        resolve(workbook);
                    });
                } else {
                    if(me.Debug==1)console.log("Contents Error.("+ext+")");
                    reject(me.makejqXHRMessage(404,"Contents Error.("+ext+")"));
                }
            }).fail(function(jqXHR, textStatus, errorThrown){
                reject({"jqXHR":jqXHR, "textStatus":textStatus});
            });
        });
    }
    ExcelDownloadItemAttach(itemid, attachment, name){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("ExcelDownloadItemAttach()");
            me.getGuid(itemid, attachment, name).then(function(value){
                if(value=="")value="0000000000000000";
                me.ExcelDownloadGuid(value).then(function(workbook){
                    resolve(workbook);
                }).catch(function(value){
                    reject(value);
                });
            }).catch(function(value){
                if(me.Debug==1)console.log(value);
                reject(value);
            });
        });
    }
    makejqXHRMessage(status, message){
        let json = {
            Id: 0,
            StatusCode: status,
            Message: message
        }
        let jsons = JSON.stringify(json)
        return({jqXHR:{
            responseJSON:json,
            responseText:jsons,
            status:status
        },textStatus:message});
    }
    getGuid(itemid, attachment, name){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("getGuid()");
            me.ApiGet(itemid, {"ApiVersion":1.0})
            .then(function(value){
                if(me.Debug==1)console.log(value);
                let RetVal = "";
                if(Object.keys(value.data.Response.Data).length == 1){
                    let ata;
                    eval(`ata = value.data.Response.Data[0].Attachments${attachment};`);
                    if(ata !== undefined){
                        ata = JSON.parse(ata);
                        if(me.Debug==1)console.log(ata);
                        for( let c = 0; c < Object.keys(ata).length; c++) {
                            if(me.Debug==1)console.log(ata[c].Name);
                            if(ata[c].Name == name){
                                RetVal = ata[c].Guid;
                            }
                        }
                    }    
                } else {
                    reject(); 
                }
                resolve(RetVal);
            }).catch(function(value){
                reject(value);
            });        
        });
    }
    LoadBlob(blob, workbook){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("LoadBlob()");
            const reader = new FileReader();
            reader.readAsArrayBuffer(blob);
            reader.onload = function(){
                workbook.xlsx.load(reader.result)
                .then(function(workbook){
                    resolve(workbook);
                });
            }    
        });
    }
    ExcelNew(sheetname){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("ExcelNew()");
            sheetname == undefined ? sheetname = "Sheet1" : sheetname = sheetname;
            if(me.Debug==1)console.log("Excel New Done.");
            const workbook = new ExcelJS.Workbook();
            workbook.addWorksheet(sheetname);
            resolve(workbook);
        });
    }
    SaveAs(workbook, filename){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("SaveAs()");
            filename == undefined ? filename = "output_"+String($p.id())+".xlsx" : filename = filename;
            if(me.Debug==1)console.log(workbook);
            workbook.xlsx.writeBuffer().then((buffer) => {
                const blob = new Blob([buffer], { 
                    "type": "application/vnd.openxmlformats-officedocument.screadsheetml.sheet" 
                });
                saveAs(blob, filename);
                resolve();
            });        
        });
    }
    ToBlob(base64, type){
        const me = this;
        if(me.Debug==1)console.log("ToBlob()");
        const bin = atob(base64.replace(/^.*,/, ''));
        const buffer = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) {
            buffer[i] = bin.charCodeAt(i);
        }
        // Blobを作成
        let blob;
        try{
            blob = new Blob([buffer.buffer], {
                "type": "application/vnd.openxmlformats-officedocument.screadsheetml.sheet" 
                //type: 'image/png'
            });
        }catch (e){
            return false;
        }
        return blob;        
    }
    getFilter(){
        var str = $("#CopyDirectUrlToClipboard").attr("onclick");
        var cmd = str.split(/(View=)(.*)('\);)/g);
        var view = JSON.parse(decodeURIComponent(cmd[2]));
        return(view);
    }
    ToExcel(recordset, filename){
        return new Promise((resolve, reject) => {
            const me = this;
            if(me.Debug==1)console.log("ToExcel()");
            filename == undefined ? filename = "" : filename = filename;
            if(me.Debug==1)console.log(recordset);
            if(Object.keys(recordset).length > 0){
                let workbook;
                me.ExcelNew().then(function(value){
                    workbook = value;
                    let sheet = workbook.worksheets[0];
                    for(let row = 0; row < Object.keys(recordset).length > 0; row++){
                        let data = recordset[row];
                        if(me.Debug==1)console.log(data);
                        if(row == 0){
                            for(let col = 0; col < Object.keys(data).length; col++){
                                //if(me.Debug==1)console.log(Object.keys(data)[col]);
                                 let r = row+1;
                                 let c = String.fromCharCode(col+65);
                                sheet.getCell(`${c}${r}`).value = Object.keys(data)[col];
                            }        
                        }
                        for(let col = 0; col < Object.keys(data).length; col++){
                            //if(me.Debug==1)console.log(Object.values(data)[col]); 
                            let r = row+2;
                            let c = String.fromCharCode(col+65);
                           sheet.getCell(`${c}${r}`).value = Object.values(data)[col];
                        }                        
                    }
                    me.SaveAs(workbook, filename);
                });

            }
        });
    }
    /*
    Sample_Promise(){
        return new Promise((resolve, reject) => {
        });
    }
    */   
}
    