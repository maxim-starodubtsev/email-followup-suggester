const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');

module.exports = (env, options) => {
  const dev = options.mode === 'development';
  
  return {
    devtool: dev ? 'inline-source-map' : 'source-map',
    entry: {
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js'],
      fallback: {
        "crypto": require.resolve("crypto-browserify"),
        "stream": require.resolve("stream-browserify"),
        "buffer": require.resolve("buffer"),
        "process": require.resolve("process/browser"),
        "vm": false
      }
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: 'ts-loader',
            options: {
              compilerOptions: {
                declaration: true,
                declarationMap: true
              }
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: 'html-loader'
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext][query]'
          }
        }
      ]
    },
    plugins: [
      new webpack.ProvidePlugin({
        Buffer: ['buffer', 'Buffer'],
        process: 'process/browser'
      }),
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['commands']
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets/*',
            to: 'assets/[name][ext][query]'
          }
        ]
      })
    ],
    devServer: {
      static: path.join(__dirname, 'dist'),
      port: 3000,
      hot: true,
      allowedHosts: 'all',
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      server: {
        type: 'https'
      },
      historyApiFallback: {
        rewrites: [
          { from: /^\/taskpane\.html$/, to: '/taskpane.html' },
          { from: /^\/commands\.html$/, to: '/commands.html' }
        ]
      },
      open: false,
      client: {
        overlay: {
          errors: true,
          warnings: false
        }
      }
    }
  };
};