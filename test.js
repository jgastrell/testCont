const str = 'S3.10.0\r\n\r\n\r\n\r\n\r\n\r\nS3.10.0';
console.log(str)

const newStr = str.replace(/(\r\n)/g, " ")
// replace(/(\r\n)/g, "")
console.log(newStr)