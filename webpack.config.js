const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  // Trusted HTTPS cert for localhost so Excel on the web can load add-in files.
  // Only needed for the local dev server — never in a production build (e.g. the
  // Cloudflare Pages build runs `npm run build` with no certs and no interactivity).
  let httpsOptions = {};
  if (dev) {
    try {
      httpsOptions = await devCerts.getHttpsServerOptions();
    } catch (e) {
      // Certs not installed yet — `npm run dev:certs` will create them.
    }
  }

  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      functions: "./src/functions/functions.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        { test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ },
        { test: /\.css$/, use: ["style-loader", "css-loader"] },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          // icons → dist/assets/
          { from: "assets", to: "assets" },
          // custom functions metadata → dist/ (manifest references /functions.json)
          { from: "assets/functions.json", to: "functions.json" },
          // custom functions runtime HTML shim → dist/
          { from: "src/functions/functions.html", to: "functions.html" },
        ],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: httpsOptions,
      },
      port: 3000,
    },
  };
};
