/* FunctionFile：清单 schema 要求 DesktopFormFactor 必须携带。
 * 本插件的实际功能全部由任务窗格（taskpane.js）执行，这里仅提供符合 schema 的空实现，
 * 供 Office 在启动时完成功能区命令注册（缺失该元素时部分 Office 版本会推迟注册）。 */
(function () {
  "use strict";
  Office.onReady(function () {});
})();
