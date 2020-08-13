# Boostrap 期末大作业

### 1830315 何伟亮

> 浏览器测试版本：google版本：83.0.4103.97，其他版本的浏览器没有测试过，可能存在一定的差异

## 提交时间：16 周周五 == 6 月 19 号

开发时间安排：

- 2020-5-28： 完成基本项目构思 + 快速原型
- 2020-5-29： 编写开发文档 +UI
- 2020-6-4： 完成首页的网页制作
- 2020-6-5~13：复习bootstrap相关知识基础
- 2020-6-13~17： 完成网页的基本制作

## 网页定位

- 静态网页

## 技术选型

- bootstrap3

- html+html5+css+css3

- JS + jQuery

- 图标效果：阿里 icon + bs3图标

- 动效效果：animate.js + Hover.js + wow.js + charts 数据表

## 要求

- 响应式

- 布局组件（3 种及以上）

  - [x] 栅格布局

  - [x] 排版
  - [x] 按钮

- CSS 组件(7 种及以上)

  > Glyphicons、字体图标、下拉菜单、按钮组、按钮式下拉菜单、输入框组、导航、导航条、路径导航、分页、标签、徽章、巨幕、 页头、缩略图、警告框、进度条、媒体对象、 列表组、面板、具有响应式特性的嵌入内容、Well

  - [x] 字体图标
  - [x] 下拉菜单
  - [x] 导航
  - [x] 输入框组
  - [x] 警告框
  - [x] 导航、导航条、路径导航
  - [x] 弹出框
  - [x] 分页
  - [x] 徽章
  - [x] 标签
  - [x] 缩略图
  - [x] 进度条
  - [x] 面板

- JS 插件 (6 种及以上)
  > 概览、过渡效果、模态框、下拉菜单、滚动监听、标签页、工具提示、弹出框、警告框、按钮、Collapse、Carousel、Affix
  - [x] 轮播图
  - [x] 标签页
  - [x] 工具提示
  - [x] 模态框
  - [x] Collapse
  - [x] Carousel
  - [x] 弹出框
  - [x] 下拉菜单

## 网站主题

Coder 程序员的网站、个人网站（编程语言相关的网站）、介绍 各种各样编程语言的特点、拒绝 996、爱好编程、流行的技术

## 需求分析

- 基于 boostrap3 开发的 Web 响应式的网页
- 后台数据模板可视化
- 滚动数字更换的效果

## 页面内容

主要色调 蓝色

页面数量 3 个

### 首页内容 `index.html`

- 主要内容

### 登陆/注册页面 `login.html`

- 登陆和注册功能

### 后台登陆页面 `admin.html`

- 数据可视化



# 网页开发文档

## 遇到的问题

### 解决问题

- logo 制作的时候想要使用svg代码绘制动画，对svg绘画的技术不熟练

  ```
  学习了好久才完成基本的效果
  ```

- 滚动视差效果的问题，图片没有报错相对固定，

  笔记：滚动视差效果的制作，前提三层div

  ```css
    // 保持固定的层，添加css属性
    width: 100%;
    height: 200px;
    background-image: url('./banner2.jpg');
    background-attachment: fixed;
  ```

- bootstrap3 默认样式的去除两种方法

  ```
  方法一：使用内联样式
  方法二：使用!important 的方式提高优先级
  ```

- bootstrap按钮的边框问题，outline属性的样式误以为是a链接的样式

  ```
  .btn:focus,
  .btn:active:focus,
  .btn.active:focus,
  .btn.focus,
  .btn:active.focus,
  .btn.active.focus {
      outline: none;          
  }
  ```

- bootstrap版本问题：bootstrap版本低于3.3.7，会使bootstrap 的某些插件无法正常使用

- bootstrap3中的徽章(不存在辅助的样式)与标签

- tooltip 插件要同时使用两种方式``data和JS``引用才能生效

  > Tooltip 插件通过data 提示工具（Tooltip）插件不像之前所讨论的下拉菜单及其他插件那样，它不是纯 CSS 插件。如需使用该插件，您必须使用 jquery 激活它（读取 javascript）。使用下面的脚本来启用页面中的所有的提示工具（tooltip）：
  > `$(function () { $("[data-toggle='tooltip']").tooltip(); });`
  > 就是说要通过dat
  > 提示工具（Tooltip）插件根据需求生成内容和标记，默认情况下是把提示工具（tooltip）放在它们的触发元素后面。您可以有以下两种方式添加提示工具（tooltip）：
  >
  > 通过 data 属性：如需添加一个提示工具（tooltip），只需向一个锚标签添加 data-toggle="tooltip" 即可。锚的 title 即为提示工具（tooltip）的文本。默认情况下，插件把提示工具（tooltip）设置在顶部。
  > 通过 JavaScript：通过 JavaScript 触发提示工具（tooltip）：
  > $('#identifier').tooltip(options)

- wow.js 依赖于 animate.compat.css 文件而不是 animate.css 文件

- 响应式的部分高度问题

- 滚动监听的使用问题等等

## 总结与心得

- boostrap3 可以很快的构建项目，但是它的一些默认样式去除比较麻烦
- 学习完bootstrap4 后感觉bootstrap3 有很多东西不合理，版本比较老旧存在问题、性能等问题
- bootstrap3 的插件比较好用，但是修改起来比较麻烦