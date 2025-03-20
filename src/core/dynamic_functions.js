// 存储所有注册的动态函数
const dynamicFunctions = {};

// 注册函数
const registerFunction = (name, func) => {
  // console.log('registerFunction', name, func);
  if (typeof func !== 'function') {
    throw new Error('注册的必须是函数');
  }
  dynamicFunctions[name] = func;
};

// 获取函数
const getFunction = (name) => {
  // console.log('getFunction', name, dynamicFunctions);
  return dynamicFunctions[name];
};

// 修改executeFunction函数以支持参数
const executeFunction = (name, args = []) => {
  // console.log('executeFunction', name, args, dynamicFunctions);
  const func = dynamicFunctions[name];
  if (!func) {
    throw new Error(`未定义的函数: ${name}`);
  }
  return func(...args);
};

// 列出所有可用函数
const listFunctions = () => {
  // console.log('listFunctions', dynamicFunctions);
  return Object.keys(dynamicFunctions);
};

// 移除函数
const removeFunction = (name) => {
  // console.log('removeFunction', name, dynamicFunctions);
  delete dynamicFunctions[name];
};

export default {
  registerFunction,
  getFunction,
  executeFunction,
  listFunctions,
  removeFunction
}; 