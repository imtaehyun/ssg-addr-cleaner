var gulp = require('gulp'),
	nodemon = require('gulp-nodemon');

gulp.task('default', function() {
	nodemon({
		script: 'parser.js',
		ext: 'js'
	})
	.on('restart', function() {
		console.log('restarted');
	});
});