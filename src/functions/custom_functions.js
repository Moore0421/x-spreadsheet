// 获取人员数据函数，支持用户ID参数
function getPeople(userId) {
  // 使用标准JavaScript语法，通过return返回数据

  
  // 如果提供了userId，可以根据ID返回不同的数据
  if (userId) {
    // 模拟根据userId返回不同用户数据
    if (userId === 10938492) {
      return {
        people: [
          { name: { firstName: '张', lastName: '三' }, age: 28, id: userId }
        ]
      };
    } else {
      return {
        people: [
          { name: { firstName: '李', lastName: '四' }, age: 32, id: userId }
        ]
      };
    }
  }
  
  // 默认返回多个用户数据
  return {
    people: [
      { name: { firstName: '张', lastName: '三' }, age: 28 },
      { name: { firstName: '李', lastName: '四' }, age: 32 }
    ]
  };
}

// 获取产品数据，也支持参数
function getProducts(categoryId) {

  
  // 根据分类ID筛选产品
  if (categoryId) {
    // 模拟分类筛选
    return {
      products: [
        { id: categoryId, name: `分类${categoryId}产品`, price: 3999 + categoryId }
      ]
    };
  }
  
  // 默认返回所有产品
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