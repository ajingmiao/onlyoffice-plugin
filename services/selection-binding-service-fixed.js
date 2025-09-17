// 修复后的 updateBinding 方法片段
// 只修改匹配逻辑部分

// 检查是否匹配rid或现有tag
var isMatch = false;

// 首先尝试直接获取控件的internalId
var controlInternalId = '';
try {
    if (typeof control.GetInternalId === 'function') {
        controlInternalId = control.GetInternalId();
    }
} catch (e) {
    console.log('获取控件internalId失败:', e);
}

console.log("控件internalId:", controlInternalId, "目标rid:", scope.rid);

// 优先使用internalId匹配（最可靠）
if (scope.rid && controlInternalId && controlInternalId === scope.rid) {
    isMatch = true;
    console.log("匹配方式: internalId精确匹配");
}
// 备选：检查tag和alias中是否包含rid（用于向后兼容）
else if (scope.rid && (tag.includes(scope.rid) || alias.includes(scope.rid))) {
    isMatch = true;
    console.log("匹配方式: tag/alias包含匹配");
}
// 尝试解析JSON格式的tag中的id字段
else if (scope.rid) {
    try {
        if (tag.startsWith("{") || tag.includes("id")) {
            var tagObj = JSON.parse(tag);
            if (tagObj && (tagObj.internalId === scope.rid || tagObj.id === scope.rid)) {
                isMatch = true;
                console.log("匹配方式: JSON tag解析匹配");
            }
        }
    } catch (jsonErr) {
        // 忽略JSON解析错误
        console.log("JSON解析失败:", jsonErr);
    }
}

// 如果没有rid匹配成功，才尝试其他方式（仅用于特殊情况）
if (!isMatch) {
    // 仅当没有rid或rid匹配失败时，才使用alias名称匹配
    if (!scope.rid && scope.displayName && alias === scope.displayName) {
        isMatch = true;
        console.log("匹配方式: alias名称匹配（无rid）");
    }
    // 注意：完全移除了tag精确匹配，因为tag不是唯一标识符
    // 多个控件可能有相同的tag值，这会导致错误匹配
}

// 调试信息
if (!isMatch) {
    console.log("❌ 控件不匹配:");
    console.log("  - controlInternalId:", controlInternalId, "vs scope.rid:", scope.rid);
    console.log("  - tag包含rid:", tag.includes(scope.rid));
    console.log("  - alias包含rid:", alias.includes(scope.rid));
}