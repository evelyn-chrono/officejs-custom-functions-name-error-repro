const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const fs = require("fs");

module.exports = (env, options) => {
  // Use mkcert or self-signed certs if available, otherwise fall back to
  // webpack-dev-server's built-in self-signed cert generation.
  let httpsOptions = true;
  const certPath = path.resolve(__dirname, "certs");
  if (
    fs.existsSync(path.join(certPath, "localhost.pem")) &&
    fs.existsSync(path.join(certPath, "localhost-key.pem"))
  ) {
    httpsOptions = {
      cert: fs.readFileSync(path.join(certPath, "localhost.pem")),
      key: fs.readFileSync(path.join(certPath, "localhost-key.pem")),
    };
  }

  return {
    entry: {
      functions: "./src/functions/functions.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["functions"],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane.html",
        chunks: [],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "./src/functions/functions.json", to: "functions.json" },
        ],
      }),
    ],
    devServer: {
      port: 3000,
      server: {
        type: "https",
        options: httpsOptions === true ? {} : httpsOptions,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  };
};
