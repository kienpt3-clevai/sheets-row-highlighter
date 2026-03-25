const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const path = require("path");
const fs = require("fs");
const os = require("os");

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.js",
    commands: "./src/commands/commands.js",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  resolve: {
    extensions: [".js"],
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: "babel-loader",
          options: {
            presets: ["@babel/preset-env"],
          },
        },
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane"],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
        { from: "manifest.xml", to: "manifest.xml" },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: {
      type: "https",
      options: {
        key: fs.readFileSync(path.join(os.homedir(), ".office-addin-dev-certs", "localhost.key")),
        cert: fs.readFileSync(path.join(os.homedir(), ".office-addin-dev-certs", "localhost.crt")),
        ca: fs.readFileSync(path.join(os.homedir(), ".office-addin-dev-certs", "ca.crt")),
      },
    },
    port: 3001,
    hot: true,
  },
};
