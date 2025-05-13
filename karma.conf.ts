"use strict";

const webpackConfig = require("./test.webpack.config.js");
const tsconfig = require("./test.tsconfig.json");
const path = require("path");

const testRecursivePath = "test/visualTest.ts";
const srcOriginalRecursivePath = "src/**/*.ts";
const coverageFolder = "coverage";

process.env.CHROME_BIN =
  require("playwright-chromium").chromium.executablePath();

module.exports = (config) => {
  config.set({
    mode: "development",
    browserNoActivityTimeout: 100000,
    browsers: ["ChromeHeadless"],
    customLaunchers: {
      ChromeDebugging: {
        base: "ChromeHeadless",
        flags: ["--remote-debugging-port=9333"],
      },
    },
    colors: true,
    frameworks: ["jasmine"],
    reporters: ["progress", "junit"],
    junitReporter: {
      outputDir: path.join(__dirname, coverageFolder),
      outputFile: "TESTS-report.xml",
      useBrowserName: false,
    },
    singleRun: true,
    plugins: [
      "karma-coverage",
      "karma-typescript",
      "karma-webpack",
      "karma-jasmine",
      "karma-sourcemap-loader",
      "karma-chrome-launcher",
      "karma-junit-reporter",
    ],
    files: [
      testRecursivePath,
      {
        pattern: srcOriginalRecursivePath,
        included: false,
        served: true,
      },
      {
        pattern: "**/*.json",
        watched: true,
        served: true,
        included: false,
      },
    ],
    preprocessors: {
      [testRecursivePath]: ["webpack", "coverage"],
    },
    typescriptPreprocessor: {
      options: tsconfig.compilerOptions,
    },
    coverageReporter: {
      dir: path.join(__dirname, coverageFolder),
      reporters: [
        // reporters not supporting the `file` property
        { type: "html", subdir: "html-report" },
        { type: "lcov", subdir: "lcov" },
        // reporters supporting the `file` property, use `subdir` to directly
        // output them in the `dir` directory
        { type: "cobertura", subdir: ".", file: "cobertura-coverage.xml" },
        { type: "lcovonly", subdir: ".", file: "report-lcovonly.txt" },
        { type: "text-summary", subdir: ".", file: "text-summary.txt" },
      ],
    },
    mime: {
      "text/x-typescript": ["ts", "tsx"],
    },
    webpack: webpackConfig,
    webpackMiddleware: {
      stats: "errors-only",
    },
  });
};
