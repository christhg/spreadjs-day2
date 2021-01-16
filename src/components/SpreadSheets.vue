<template>
  <div class="home">
    <div class="toolbar">
        <input type="file" class="el-button" @change="importExcel($event)" style="height: 40px;"/>
        <el-button @click="exportExcel()">导入 Excel</el-button>  
        <el-button @click="exportExcel()">导出 Excel</el-button>
      <el-button-group style="margin-left:10px">
        <el-button type="primary" icon="el-icon-bottom" @click="loadTemplate()">加载</el-button>
        <el-button type="primary" icon="el-icon-top" @click="updateTemplate()">更新</el-button>
      </el-button-group>

      <!-- <el-button-group style="margin-left:10px">
        <el-button type="primary" icon="el-icon-bottom" @click="download()"></el-button>
        <el-button type="primary" icon="el-icon-top" @click="upload()"></el-button>
      </el-button-group> -->
    </div>
    <div class="spreadWrapper">
      <div ref="formulaBar" class="formulaBar" contenteditable="true" spellcheck="false"></div>
      <gc-spread-sheets class="spreadHost" @workbookInitialized="workbookInitialized($event)">
        <gc-worksheet></gc-worksheet>
      </gc-spread-sheets>
      <div ref="statusBar" class="statusBar"></div>
    </div>
  </div>
</template>

<script lang="ts">
// import Vue from "vue";
import { Component, Vue, Watch } from 'vue-property-decorator';
// import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css";
import "@grapecity/spread-sheets-vue";
import * as GC from "@grapecity/spread-sheets";
import * as ExcelIO from "@grapecity/spread-excelio";
import "@grapecity/spread-sheets-barcode";
import "@grapecity/spread-sheets-charts";
import "@grapecity/spread-sheets-shapes";
import "@grapecity/spread-sheets-languagepackages";
import "@grapecity/spread-sheets-print";
import "@grapecity/spread-sheets-pdf";
import "@grapecity/spread-sheets-pivot-addon";
import "@grapecity/spread-sheets-resources-zh";

import FileSaver from "file-saver";
import httpUtils from "@/HttpUtils";
import { Notification } from 'element-ui';

//GC.Spread.Sheets.LicenseKey = ExcelIO.LicenseKey = ""

GC.Spread.Common.CultureManager.culture("zh-cn");

// export default Vue.extend({});
@Component
export default class SpreadSheets extends Vue {
  spread: GC.Spread.Sheets.Workbook | null;
  constructor() {
    super();
    this.spread = null;
  }
  @Watch('$route',{ immediate: true })
  private changeRouter(route: any){
      // console.log(route)
      if(route.params.id){
        this.download(route.params.id)
      }
  }
  workbookInitialized(spread: GC.Spread.Sheets.Workbook) {
    this.spread = spread;

    let formulaBar = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(this.$refs.formulaBar as HTMLElement, {} as GC.Spread.Sheets.FormulaTextBox.IFormulaTextBoxOptions);
    formulaBar.workbook(this.spread);
    this.spread.focus();
    
    let statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(this.$refs.statusBar as HTMLElement);
    statusBar.bind(spread);

    this.registEvent(this.spread);
  }
  sendMessage(info:any){

      let cellRef = GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(info.row, info.col, info.rowCount|1, info.colCount|1), 0, 0, GC.Spread.Sheets.CalcEngine.RangeReferenceRelative.allRelative, false);
      let message = "单元格" + cellRef + "发生了变化！";

      Notification({title: "同步", message: message, type:"info"});
  }
  registEvent(spread: GC.Spread.Sheets.Workbook){
    let self = this;
    spread.bind(GC.Spread.Sheets.Events.ValueChanged, (s:any, e:any) => {
      self.sendMessage(e);
    })
    spread.bind(GC.Spread.Sheets.Events.RangeChanged, (s:any, e:any) => {
      self.sendMessage(e);
    })
  }
  importExcel(event: any) {
    let file = event.target.files[0];
    let self = this;

    let excelIO = new ExcelIO.IO();
    excelIO.open(file, (spreadJSON: object) => {
      if (self.spread) {
        self.spread.fromJSON(spreadJSON);
      }
    });
  }
  exportExcel() {
    let self = this;
    if (self.spread) {
      let spreadJSON = JSON.stringify(self.spread.toJSON());
      let excelIO = new ExcelIO.IO();
      excelIO.save(spreadJSON, (blob:Blob)=>{
        FileSaver.saveAs(blob, "export.xlsx");
      })
    }
  }

  loadTemplate() {
    let self = this;
    let formData = new FormData();
    formData.append("fileName", "path")
    httpUtils.post("spread/loadTemplate", formData ,{responseType: "json"}).then(response=>{
      if(response){
          if (self.spread) {
            self.spread.fromJSON(response as object);
            Notification({title: "下载成功", message: "", type:'success'});
          }
      }
    });
  }

  updateTemplate() {
    let self = this;
    if (self.spread) {
      let spreadJSON = JSON.stringify(self.spread.toJSON());
        let formData = new FormData();
        formData.append("jsonString", spreadJSON);
        formData.append("fileName", "fileName");
          httpUtils.post("spread/updateTemplate", formData).then(response=>{
            if(response){
              Notification({title: response as string, message: "", type:'success'});
            }
          });
    }
  }


  download(id:any) {
    let self = this;
    let formData = new FormData();
    formData.append("path", "path")
    httpUtils.post("spread/getExcel", formData ,{responseType: "blob"}).then(response=>{
      if(response){
        let excelIO = new ExcelIO.IO();
        excelIO.open(response as Blob, (spreadJSON: object) => {
          if (self.spread) {
            self.spread.fromJSON(spreadJSON);
            Notification({title: "下载成功", message: "", type:'success'});
          }
        });
      }
    });
  }

  upload() {
    let self = this;
    if (self.spread) {
      let spreadJSON = JSON.stringify(self.spread.toJSON());
      let excelIO = new ExcelIO.IO();
      excelIO.save(spreadJSON, (blob:Blob)=>{
        let formData = new FormData();
        formData.append("excelFile", blob);
          httpUtils.post("spread/saveExcel", formData).then(response=>{
            if(response){
              Notification({title: response as string, message: "", type:'success'});
            }
          });
      })
    }
  }
}
</script>

<style>
.home,
.spreadWrapper {
  height: calc(100% - 45px);
}
.spreadHost {
  height: calc(100% - 75px);
  width: 100%;
}
.toolbar {
  padding-bottom: 5px;
}
.formulaBar{
  height: 43px;
  width: calc(100% - 3px);
  margin-bottom: 2px;
  border: 1px solid #808080;
  background: white;
}
.statusBar{
  height: 25px;
  width: 100%;
}
</style>