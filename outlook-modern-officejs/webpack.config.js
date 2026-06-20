/* eslint-disable no-undef */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

const urlDev = "https://localhost:3000/";

async function getHttpsOptions() {
  try {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
  } catch {
    // Si encara no s'han instal·lat els certificats de dev, retornem buit.
    return {};
  }
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  const config = {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane.ts",
      commands: "./src/commands.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "src/taskpane.css", to: "taskpane.css" },
          { from: "assets", to: "assets", noErrorOnMissing: true },
          { from: "manifest.xml", to: "manifest.xml", noErrorOnMissing: true },
        ],
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options: env && env.WEBPACK_BUILD ? {} : await getHttpsOptions(),
      },
      port: 3000,
    },
  };

  return config;
};

// Exposat per a referència; urlDev s'usa al manifest.
module.exports.urlDev = urlDev;
