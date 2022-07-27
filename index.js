const express = require("express");
const cors = require("cors");
const multer = require("multer");
const app = express();
const path = require("path");
const fs = require("fs");
const uuid = require("uuid");
const json2xls = require("json2xls");
const { ExcToJSON } = require("./helpers");

app.use(function (req, res, next) {
  res.header("Access-Control-Allow-Origin", "*");
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept"
  );
  next();
});

app.use(cors({ origin: "*" }));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.set("view engine", "ejs");

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, path.join(__dirname, "uploads"));
  },
  filename: function (req, file, cb) {
    const extn = file.originalname.split(".").reverse()[0];
    cb(null, `${uuid.v4()}.${extn}`);
  },
});

const upload = multer({ storage: storage, dest: "/uploads" });

app.get("/", function (req, res) {
  res.render("pages/index");
});

app.post(
  "/upload",
  upload.fields([
    { name: "master", maxCount: 1 },
    { name: "append", maxCount: 10 },
  ]),
  (req, res) => {
    try {
      if (Object.keys(req.files)?.length !== 2) {
        return res.redirect("/");
      }

      const masterFilePath = req.files.master[0].path;
      const appendFilesPathArr = req.files.append.map((file) => file.path);

      const files = [masterFilePath, appendFilesPathArr].flat();
      const arr = [];
      for (let index = 0; index < files.length; index++) {
        const obj = ExcToJSON(files[index]);
        arr.push(obj);
      }
      const obj = {};
      arr.map((item) => {
        for (let key in item) {
          obj[key] = obj[key] ? [...obj[key], ...item[key]] : [...item[key]];
        }
      });

      const xls = json2xls(Object.values(obj).flat());

      const finalFile = path.join(__dirname, "result", `final.xlsx`);
      fs.writeFileSync(finalFile, xls, "binary");

      res.sendFile(finalFile);
    } catch (error) {
      console.error(error.message);
      res.redirect("/error");
    } finally {
      // clean up
      setTimeout(() => {
        fs.rmSync(path.join(__dirname, "result"), {
          recursive: true,
          force: true,
        });
        fs.rmSync(path.join(__dirname, "uploads"), {
          recursive: true,
          force: true,
        });
        fs.mkdirSync(path.join(__dirname, "result"));
        fs.mkdirSync(path.join(__dirname, "uploads"));
      }, 100);
    }
  }
);

app.get("/error", (req, res) => {
  res.render("pages/500");
});

app.get("/*", (req, res) => {
  res.render("pages/404");
});

app.listen(8000, () => {
  console.log("Server running on port 8000");
});
