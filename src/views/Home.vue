ㄝ<template>
  <div class="home">
    <!-- <img alt="Vue logo" src="../assets/logo.png">
    <HelloWorld msg="Welcome to Your Vue.js App"/>-->

    <el-container>
      <el-header>
          表格文档在线编辑系统
          <router-link to="/">Home</router-link> |
          <router-link to="/editor">Editor</router-link>
      </el-header>
       <el-header>
          表格文档在线编辑系统
          <router-link to="/">功能1</router-link> |
          <router-link to="/editor">功能2</router-link>
          <router-link to="/editor">功能3</router-link>
      </el-header>     
      <el-container>
        <!-- <el-aside width="200px">Aside</el-aside> -->
        <el-container>
          <el-main padding="5px">
              <SpreadSheets />
          </el-main>
          <el-footer>Footer</el-footer>
        </el-container>
      </el-container>
    </el-container>
  </div>
</template>

<script>
// @ is an alias to /src
import HelloWorld from "@/components/HelloWorld.vue";
import SpreadSheets from "@/components/SpreadSheets";

export default {
  name: "Home",
  components: {
    HelloWorld,
    SpreadSheets
  }
};
</script>

<style>
  .el-header, .el-footer {
    background-color: #B3C0D1;
    color: #333;
    text-align: center;
    line-height: 60px;
  }
  
  .el-aside {
    background-color: #D3DCE6;
    color: #333;
    text-align: center;
    line-height: 200px;
  }
  
  .el-main {
    background-color: #E9EEF3;
    color: #333;
    /* text-align: center;
    line-height: 160px; */
  }
  
  body > .el-container {
    margin-bottom: 40px;
  }
  
  .el-container:nth-child(5) .el-aside,
  .el-container:nth-child(6) .el-aside {
    line-height: 260px;
  }
  
  .el-container:nth-child(7) .el-aside {
    line-height: 320px;
  }
  
  html,body,#app,.home,.el-container,.el-main{
    height: 100%;
    margin: 0;
  }

</style>