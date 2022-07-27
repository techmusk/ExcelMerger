"use strict";

const excelToJson = require("convert-excel-to-json");

const fs = require("fs");

const moment = require("moment");
/**
 * @description
 * In case all objects in an array of objects doesn't have some missing keys then this function will append those missing keys in all other objects and make it equal
 */


function handleMakeEqual(arr) {
  const keys = arr.reduce((acc, curr) => (Object.keys(curr).forEach(key => acc.add(key)), acc), new Set());
  const output = arr.map(item => [...keys].reduce((acc, key) => {
    var _item$key;

    return acc[key] = (_item$key = item[key]) !== null && _item$key !== void 0 ? _item$key : "", acc;
  }, {}));
  return output;
}

function handleDigit(num) {
  return String(num).length === 1 ? `0${num}` : num;
}

function filter(data = []) {
  const arr = handleMakeEqual(data);
  const header = arr[0];
  const body = arr.slice(1);
  const raw = body.map(item => {
    const instance = {};

    for (const key in item) {
      var _instance$newInstance;

      let newInstanceKey = header[key];

      if (newInstanceKey in instance) {
        // already another key exists in instance
        const reg = new RegExp(newInstanceKey, "gi");
        const noOfOccurence = Object.keys(instance).filter(item => reg.test(item)).length;
        newInstanceKey = `${newInstanceKey}-${noOfOccurence}`;
        instance[newInstanceKey] = item[key];
      } // manipulating date


      if (/Invoice Date/gi.test(newInstanceKey) || /Cancellation Date/gi.test(newInstanceKey)) {
        // date validation
        if (!item[key]) {
          instance[newInstanceKey] = "";
        } else if (new Date(item[key]) == "Invalid Date") {
          // 14/06/2022 10:46:55 AM
          let [day, month, year] = item[key].split(" ")[0].split("/");
          instance[newInstanceKey] = moment(new Date(year, parseInt(month) - 1, day)).format("DD-MMM-YY");
        } else {
          instance[newInstanceKey] = moment(item[key]).format("DD-MMM-YY");
        }
      } else if (item[key] || item[key] === 0) {
        // item[key] has a zero as value or some value
        instance[newInstanceKey] = item[key];
      } else {
        // item[key] is null or undefined
        instance[newInstanceKey] = "";
      } // manipulating voucher type key


      if (instance["Invoice Number"] && // Invoice Number
      instance["Invoice Number"].trim().length > 0 && /Voucher Type/gi.test(newInstanceKey) && (!instance[newInstanceKey] || ((_instance$newInstance = instance[newInstanceKey]) === null || _instance$newInstance === void 0 ? void 0 : _instance$newInstance.trim().length) === 0)) {
        instance[newInstanceKey] = `${instance["Invoice Number"].substring(0, 3)} INVOICE`;
      }
 
      if (/Customer Name/gi.test(newInstanceKey) && (!instance["Customer Name"] || instance["Customer Name"].trim().length === 0)) {
        instance["Customer Name"] = "Cash";
      }
    }

    return instance;
  });
  return raw;
}

function ExcToJSON(excelFilePath) {
  const obj = {};
  const result = excelToJson({
    source: fs.readFileSync(excelFilePath)
  });

  for (const key in result) {
    const json = filter(result[key]);
    obj[key] = json;
  }

  return obj;
}

module.exports = {
  ExcToJSON
};