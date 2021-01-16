<template>
  <div class="home">
    <div class="designerWrapper">
      <gc-spread-sheets-designer
        :styleInfo="styleInfo"
        :config="config"
        @designerInitialized="designerInitialized"
      >
      </gc-spread-sheets-designer>
    </div>
  </div>
</template>

<script lang="ts">
// import Vue from "vue";
import { Component, Vue, Watch } from "vue-property-decorator";

import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css";
import "@grapecity/spread-sheets-designer/styles/gc.spread.sheets.designer.min.css";
import "@grapecity/spread-sheets-designer-resources-cn";
import * as GC from "@grapecity/spread-sheets";
import * as ExcelIO from "@grapecity/spread-excelio";
import "@grapecity/spread-sheets-resources-zh";
import * as DesignerGC from "@grapecity/spread-sheets-designer";
import "@grapecity/spread-sheets-designer-vue";

GC.Spread.Common.CultureManager.culture("zh-cn");

import httpUtils from "@/HttpUtils";
import { Notification } from "element-ui";

//GC.Spread.Sheets.LicenseKey = ExcelIO.LicenseKey = ""
// DesignerGC.Spread.Sheets.Designer.LicenseKey = "";

@Component
export default class Designer extends Vue {
  designer: DesignerGC.Spread.Sheets.Designer.Designer | null;
  styleInfo: object;
  config?: object;
  constructor() {
    super();
    this.designer = null;
    this.styleInfo = { height: "100%", width: "100%" }; //設計器寬高
    this.config = this.getCustomerConfig(); //設計器配置檔
  }

  //獲取配置檔
  getCustomerConfig() {
    let customerConfig = JSON.parse(
      JSON.stringify(DesignerGC.Spread.Sheets.Designer.DefaultConfig)
    );
    var customerRibbon = {
      id: "operate",
      text: "操作",
      buttonGroups: [
        {
          label: "文件操作",
          thumbnailClass: "ribbon-thumbnail-spreadsettings",
          commandGroup: {
            children: [
              {
                direction: "vertical",
                commands: ["loadTemplate", "updateTemplate"],
              },
              {
                direction: "vertical",
                commands: ["loadTemplate"],
              },
              {
                type: "separator",
              },
              {
                direction: "vertical",
                commands: ["updateTemplate"],
              },
            ],
          },
        },
      ],
    };

    let loadTemplateCommand = {
      iconClass: "ribbon-button-download",
      text: "加载",
      // bigButton: true,
      commandName: "loadTemplate",
      execute: this.loadTemplate,
    };
    let updateTemplateCommand = {
      iconClass: "ribbon-button-upload",
      text: "更新",
      // bigButton: true,
      commandName: "updateTemplate",
      execute: this.updateTemplate,
    };

    customerConfig.ribbon.push(customerRibbon);
    customerConfig.commandMap = {
      loadTemplate: loadTemplateCommand,
      updateTemplate: updateTemplateCommand,
    };
    console.log(customerConfig);
    // customerConfig.fileMenu = undefined;
    return customerConfig;
  }

  //設計器初始化
  designerInitialized(designer: DesignerGC.Spread.Sheets.Designer.Designer) {
    this.designer = designer;
    // this.loadTemplate(designer)
  }

  async loadTemplate(designer: DesignerGC.Spread.Sheets.Designer.Designer) {
    let spread = designer.getWorkbook() as GC.Spread.Sheets.Workbook;
    let formData = new FormData();
    formData.append("fileName", "path");
    let templateJSON = await httpUtils.post("spread/loadTemplate", formData, {
      responseType: "json",
    });
    spread.fromJSON(templateJSON as object);
    Notification({ title: "下载成功", message: "", type: "success" });
  }

  async updateTemplate(designer: DesignerGC.Spread.Sheets.Designer.Designer) {
    let spread = designer.getWorkbook() as GC.Spread.Sheets.Workbook;
    let spreadJSON = JSON.stringify(spread.toJSON());
    let formData = new FormData();
    formData.append("jsonString", spreadJSON);
    formData.append("fileName", "fileName");
    let result = await httpUtils.post("spread/updateTemplate", formData);
    Notification({
      title: result as string,
      message: "",
      type: "success",
    });
  }
}
</script>

<style>
.home,
.designerWrapper {
  height: calc(100% - 45px);
}
.ribbon-button-download {
  background-image: url("data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iaXNvLTg4NTktMSI/Pg0KPCEtLSBHZW5lcmF0b3I6IEFkb2JlIElsbHVzdHJhdG9yIDE5LjAuMCwgU1ZHIEV4cG9ydCBQbHVnLUluIC4gU1ZHIFZlcnNpb246IDYuMDAgQnVpbGQgMCkgIC0tPg0KPHN2ZyB2ZXJzaW9uPSIxLjEiIGlkPSJDYXBhXzEiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIHg9IjBweCIgeT0iMHB4Ig0KCSB2aWV3Qm94PSIwIDAgNTEyLjI5MyA1MTIuMjkzIiBzdHlsZT0iZW5hYmxlLWJhY2tncm91bmQ6bmV3IDAgMCA1MTIuMjkzIDUxMi4yOTM7IiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCjxwYXRoIHN0eWxlPSJmaWxsOiNCQkRFRkI7IiBkPSJNNDAyLjE0OCwxNDkuNjA2QzM4NC4zMzgsNjMuMDU0LDI5OS43MzUsNy4zMjgsMjEzLjE4MywyNS4xMzgNCglDMTM5LjA3LDQwLjM4OSw4NS43NzQsMTA1LjQ3Miw4NS40MzQsMTgxLjEzNmMwLDMuNjA1LDAuMTQ5LDcuMjk2LDAuNDY5LDExLjJDMzMuMTc4LDE5Ny45MTctNS4wNCwyNDUuMTgzLDAuNTQxLDI5Ny45MDgNCgljNS4xNzMsNDguODcsNDYuNDE2LDg1Ljk0Myw5NS41NTksODUuODk1aDExLjJjLTAuMjU2LTMuNTQxLTAuNTMzLTcuMDYxLTAuNTMzLTEwLjY2N2MwLTc2LjU4Myw2Mi4wODMtMTM4LjY2NywxMzguNjY3LTEzOC42NjcNCglTMzg0LjEsMjk2LjU1MywzODQuMSwzNzMuMTM2YzAsMy42MDUtMC4yNzcsNy4xMjUtMC41MzMsMTAuNjY3aDExLjJjNjQuNzMsMC4xNzcsMTE3LjM0OC01Mi4xNTQsMTE3LjUyNS0xMTYuODg1DQoJQzUxMi40NjIsMjA0LjgwNyw0NjQuMTQ4LDE1My4zNDgsNDAyLjE0OCwxNDkuNjA2TDQwMi4xNDgsMTQ5LjYwNnoiLz4NCjxjaXJjbGUgc3R5bGU9ImZpbGw6IzRDQUY1MDsiIGN4PSIyNDUuNDM0IiBjeT0iMzczLjEzNiIgcj0iMTE3LjMzMyIvPg0KPGc+DQoJPHBhdGggc3R5bGU9ImZpbGw6I0ZBRkFGQTsiIGQ9Ik0yNDUuNDM0LDQ0Ny44MDNjLTUuODkxLDAtMTAuNjY3LTQuNzc2LTEwLjY2Ny0xMC42Njd2LTEyOGMwLTUuODkxLDQuNzc2LTEwLjY2NywxMC42NjctMTAuNjY3DQoJCXMxMC42NjcsNC43NzYsMTAuNjY3LDEwLjY2N3YxMjhDMjU2LjEsNDQzLjAyNywyNTEuMzI1LDQ0Ny44MDMsMjQ1LjQzNCw0NDcuODAzeiIvPg0KCTxwYXRoIHN0eWxlPSJmaWxsOiNGQUZBRkE7IiBkPSJNMjQ1LjQzNCw0NDcuODAzYy0yLjgzMSwwLjAwNS01LjU0OC0xLjExNS03LjU1Mi0zLjExNWwtNDIuNjY3LTQyLjY2Nw0KCQljLTQuMDkzLTQuMjM3LTMuOTc1LTEwLjk5LDAuMjYyLTE1LjA4M2M0LjEzNC0zLjk5MywxMC42ODctMy45OTMsMTQuODIxLDBsMzUuMTM2LDM1LjExNWwzNS4xMTUtMzUuMTE1DQoJCWM0LjIzNy00LjA5MywxMC45OS0zLjk3NSwxNS4wODMsMC4yNjJjMy45OTMsNC4xMzQsMy45OTMsMTAuNjg3LDAsMTQuODIxbC00Mi42NjcsNDIuNjY3DQoJCUMyNTAuOTY1LDQ0Ni42ODIsMjQ4LjI1Nyw0NDcuODAyLDI0NS40MzQsNDQ3LjgwM3oiLz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjwvc3ZnPg0K");
}
.ribbon-button-upload {
  background-image: url("data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iaXNvLTg4NTktMSI/Pg0KPCEtLSBHZW5lcmF0b3I6IEFkb2JlIElsbHVzdHJhdG9yIDE5LjAuMCwgU1ZHIEV4cG9ydCBQbHVnLUluIC4gU1ZHIFZlcnNpb246IDYuMDAgQnVpbGQgMCkgIC0tPg0KPHN2ZyB2ZXJzaW9uPSIxLjEiIGlkPSJDYXBhXzEiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIHg9IjBweCIgeT0iMHB4Ig0KCSB2aWV3Qm94PSIwIDAgNTEyLjI5MyA1MTIuMjkzIiBzdHlsZT0iZW5hYmxlLWJhY2tncm91bmQ6bmV3IDAgMCA1MTIuMjkzIDUxMi4yOTM7IiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCjxwYXRoIHN0eWxlPSJmaWxsOiNCQkRFRkI7IiBkPSJNNDAyLjE0OCwxNDkuNjA2QzM4NC4zMzgsNjMuMDU0LDI5OS43MzUsNy4zMjgsMjEzLjE4MywyNS4xMzgNCglDMTM5LjA3LDQwLjM4OSw4NS43NzQsMTA1LjQ3Miw4NS40MzQsMTgxLjEzNmMwLDMuNjA1LDAuMTQ5LDcuMjk2LDAuNDY5LDExLjJDMzMuMTc4LDE5Ny45MTctNS4wNCwyNDUuMTgzLDAuNTQxLDI5Ny45MDgNCgljNS4xNzMsNDguODcsNDYuNDE2LDg1Ljk0Myw5NS41NTksODUuODk1aDExLjJjLTAuMjU2LTMuNTQxLTAuNTMzLTcuMDYxLTAuNTMzLTEwLjY2N2MwLTc2LjU4Myw2Mi4wODMtMTM4LjY2NywxMzguNjY3LTEzOC42NjcNCglTMzg0LjEsMjk2LjU1MywzODQuMSwzNzMuMTM2YzAsMy42MDUtMC4yNzcsNy4xMjUtMC41MzMsMTAuNjY3aDExLjJjNjQuNzMsMC4xNzcsMTE3LjM0OC01Mi4xNTQsMTE3LjUyNS0xMTYuODg1DQoJQzUxMi40NjIsMjA0LjgwNyw0NjQuMTQ4LDE1My4zNDgsNDAyLjE0OCwxNDkuNjA2TDQwMi4xNDgsMTQ5LjYwNnoiLz4NCjxjaXJjbGUgc3R5bGU9ImZpbGw6IzRDQUY1MDsiIGN4PSIyNDUuNDM0IiBjeT0iMzczLjEzNiIgcj0iMTE3LjMzMyIvPg0KPGc+DQoJPHBhdGggc3R5bGU9ImZpbGw6I0ZBRkFGQTsiIGQ9Ik0yNDUuNDM0LDQ0Ny44MDNjLTUuODkxLDAtMTAuNjY3LTQuNzc2LTEwLjY2Ny0xMC42Njd2LTEyOGMwLTUuODkxLDQuNzc2LTEwLjY2NywxMC42NjctMTAuNjY3DQoJCXMxMC42NjcsNC43NzYsMTAuNjY3LDEwLjY2N3YxMjhDMjU2LjEsNDQzLjAyNywyNTEuMzI1LDQ0Ny44MDMsMjQ1LjQzNCw0NDcuODAzeiIvPg0KCTxwYXRoIHN0eWxlPSJmaWxsOiNGQUZBRkE7IiBkPSJNMjg4LjEsMzYyLjQ3Yy0yLjgzMSwwLjAwNS01LjU0OC0xLjExNS03LjU1Mi0zLjExNWwtMzUuMTE1LTM1LjEzNmwtMzUuMTE1LDM1LjEzNg0KCQljLTQuMjM3LDQuMDkzLTEwLjk5LDMuOTc1LTE1LjA4My0wLjI2MmMtMy45OTMtNC4xMzQtMy45OTMtMTAuNjg3LDAtMTQuODIxbDQyLjY2Ny00Mi42NjdjNC4xNjUtNC4xNjQsMTAuOTE3LTQuMTY0LDE1LjA4MywwDQoJCWw0Mi42NjcsNDIuNjY3YzQuMTU5LDQuMTcyLDQuMTQ5LDEwLjkyNi0wLjAyNCwxNS4wODVDMjkzLjYzLDM2MS4zNSwyOTAuOTIzLDM2Mi40NjksMjg4LjEsMzYyLjQ3eiIvPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPGc+DQo8L2c+DQo8Zz4NCjwvZz4NCjxnPg0KPC9nPg0KPC9zdmc+DQo=");
}
</style>