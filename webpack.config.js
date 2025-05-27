const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = (env, argv) => {
  const config = {
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/index.ts"
    },
    output: {
      filename: "[name].js",
      path: path.resolve(__dirname, "dist/app"),
      clean: true,
      publicPath: "./"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx"],
      alias: {
        "@": path.resolve(__dirname, "src")
      }
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: [
            {
              loader: "ts-loader",
              options: {
                transpileOnly: true          // 👉 skip type-checking at webpack time
              }
            }
          ]
        },
        {
          test: /\.css$/i,
          use: ["style-loader", "css-loader"]
        },
        {
          test: /pdf\.worker\.min\.js$/,
          type: "asset/resource",
          generator: {
            filename: "pdf.worker.min.js"
          }
        },
        {
          test: /\.svg$/,
          type: "asset/inline"
        },
        {
          test: /\.(png|jpe?g|gif)$/i,
          type: "asset/resource"
        },
        {
          test: /\.(woff|woff2|eot|ttf|otf)$/i,
          type: "asset/resource"
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["taskpane"]
      }),
      new HtmlWebpackPlugin({
        template: "./src/commands/commands.html",
        filename: "commands.html",
        chunks: ["commands"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: path.resolve(__dirname, "node_modules/pdfjs-dist/build/pdf.worker.min.js"),
            to: "pdf.worker.min.js"
          }
        ]
      })
    ],
    optimization: {
      splitChunks: {
        chunks: "all"
      }
    }
  };

  if (argv.mode === "development") {
    config.devtool = "eval-source-map";
    config.devServer = {
      static: {
        directory: path.join(__dirname, "dist/app")
      },
      compress: true,
      port: 3000,
      host: "localhost",
      https: true,
      hot: true,
      open: false,
      allowedHosts: "all",
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      client: {
        overlay: {
          errors: true,
          warnings: false
        }
      }
    };
  }

  return config;
};