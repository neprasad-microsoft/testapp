module.exports = function override(config, env) {
    if (env === 'production') {
        console.log('Disabling minification for the production build...');
        // Disable optimization.minimize
        config.optimization.minimize = false;
        // If you have more specific configurations, you can use the following lines
        // to disable them individually:
        // config.optimization.minimizer = [];
        // config.optimization.splitChunks = { chunks: 'all', minSize: 30000, maxSize: 0 };
        // config.optimization.runtimeChunk = false;
    }
    return config;
};
