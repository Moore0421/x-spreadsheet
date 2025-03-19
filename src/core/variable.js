import dynamicFunctions from './dynamic_functions';

export default class Variable {
  constructor(rootContext) {
    this.rootContext = rootContext;
    this.map = {};
  }

  setVariableMap(map) {
    this.map = map;
    this.rootContext.sheet.table.render();
  }

  getVariableMap() {
    return this.map;
  }

  storeVariable(text, initialValue) {
    this.map[text] = initialValue;
  }

  removeVariable(text) {
    delete this.map[text];
  }

  // 注册动态函数
  registerFunction(name, func) {
    dynamicFunctions.registerFunction(name, func);
    this.rootContext.sheet.table.render();
  }
  
  // 注销动态函数
  unregisterFunction(name) {
    dynamicFunctions.removeFunction(name);
    this.rootContext.sheet.table.render();
  }
  
  // 获取所有注册的动态函数
  getFunctions() {
    return dynamicFunctions.listFunctions();
  }
}
