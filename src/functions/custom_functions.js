// 获取人员数据函数
function getPeople() {
  // 使用标准JavaScript语法，通过return返回数据
  return {
    people: [
      { name: { firstName: '张', lastName: '三' }, age: 28 },
      { name: { firstName: '李', lastName: '四' }, age: 32 }
    ]
  };
}

// 获取产品数据
function getProducts() {
  return {
    products: [
      { id: 1, name: '手机', price: 3999 },
      { id: 2, name: '平板', price: 2599 },
      { id: 3, name: '电脑', price: 6999 }
    ]
  };
}

// 导出所有自定义函数
export {
  getPeople,
  getProducts
}; 