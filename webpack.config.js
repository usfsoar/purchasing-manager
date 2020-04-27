const GasPlugin = require("gas-webpack-plugin");

module.exports = {
  mode: "production",
  entry: "./build",
  output: {
    path: __dirname,
    filename: "Code.js",
    libraryTarget: "var",
  },
  plugins: [new GasPlugin()],
  devtool: false,
};
