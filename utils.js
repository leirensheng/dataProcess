function padLeftZero(str) {
  return ('00' + str).substr(str.length)
}
module.exports = {
  getNameLength(name) {
    let length = 0
    for (let i = 0; i < name.length; i++) {
      const isChinese = /[\u4e00-\u9fa5]|（|）|；|，|。|【|】/.test(name[i])
      length += isChinese ? 2 : 1
    }
    return length
  },
  formatDate(date, fmt = "yyyy-MM-dd") {
    if (typeof date !== "object") {
      date = new Date(Number(date));
    }
    if (/(y+)/.test(fmt)) {
      fmt = fmt.replace(
        RegExp.$1,
        (date.getFullYear() + "").substr(4 - RegExp.$1.length)
      );
    }
    const o = {
      "M+": date.getMonth() + 1,
      "d+": date.getDate(),
      "h+": date.getHours(),
      "m+": date.getMinutes(),
      "s+": date.getSeconds(),
    };
    for (const k in o) {
      if (new RegExp(`(${k})`).test(fmt)) {
        const str = o[k] + "";
        fmt = fmt.replace(
          RegExp.$1,
          RegExp.$1.length === 1 ? str : padLeftZero(str)
        );
      }
    }
    return fmt;
  },
};
