/* eslint-disable no-undef */

const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const Dotenv = require("dotenv-webpack");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const isDevelopment = options.mode === "development";
  const isStaging = Boolean(env && env.staging);
  const buildTarget = isDevelopment ? "development" : isStaging ? "staging" : "production";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      dialog: ["./src/taskpane/fallback/fallbackauthdialog.ts"],
    },
    output: {
      path: path.resolve(__dirname, "dist", buildTarget),
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    optimization: {
      nodeEnv: false,
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "auth.html",
        template: "./src/taskpane/fallback/auth.html",
        chunks: [],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/taskpane/fallback/dialog.html",
        chunks: ["dialog"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: isDevelopment
              ? "manifest-dev.xml"
              : isStaging
                ? "manifest-stag.xml"
                : "manifest-prod.xml",
            to: "manifest.xml",
          },
        ],
      }),
      new Dotenv({
        path: isDevelopment
          ? ".env.development"
          : isStaging
            ? ".env.staging"
            : ".env.production",
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env && env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
