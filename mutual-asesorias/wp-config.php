<?php
/**
 * The base configurations of the WordPress.
 *
 * This file has the following configurations: MySQL settings, Table Prefix,
 * Secret Keys, and ABSPATH. You can find more information by visiting
 * {@link https://codex.wordpress.org/Editing_wp-config.php Editing wp-config.php}
 * Codex page. You can get the MySQL settings from your web host.
 *
 * This file is used by the wp-config.php creation script during the
 * installation. You don't have to use the web site, you can just copy this file
 * to "wp-config.php" and fill in the values.
 *
 * @package WordPress
 */

// ** MySQL settings - You can get this info from your web host ** //
/** The name of the database for WordPress */
define('DB_NAME', 'mutual_asesorias');

/** MySQL database username */
define('DB_USER', 'root');

/** MySQL database password */
define('DB_PASSWORD', 'root');

/** MySQL hostname */
define('DB_HOST', 'localhost');

/** Database Charset to use in creating database tables. */
define('DB_CHARSET', 'utf8mb4');

/** The Database Collate type. Don't change this if in doubt. */
define('DB_COLLATE', '');

/**#@+
 * Authentication Unique Keys and Salts.
 *
 * Change these to different unique phrases!
 * You can generate these using the {@link https://api.wordpress.org/secret-key/1.1/salt/ WordPress.org secret-key service}
 * You can change these at any point in time to invalidate all existing cookies. This will force all users to have to log in again.
 *
 * @since 2.6.0
 */
define('AUTH_KEY',         'q:n8^E_ E((4-LSo&7Cr)oYN/eQp0,Ib6k47mxfR:qkx56gci6O9Q+8}-Io +/r$');
define('SECURE_AUTH_KEY',  'X}3h7/lpN3/5:GuPy<S(w/|+d Clu1EE%-J$d_iVU*Y,zhKM[26-ED.ns -sq`-E');
define('LOGGED_IN_KEY',    '=CWM(VuHtzC#jOnn*ONHjETr,(:x.{-ng-2J`]>N-8OgsB w(I<hL;G2ZR/YY#OC');
define('NONCE_KEY',        '];(k$+$Y!xvHflW4|zF59}*Hj;jM5SAtAdniOE#JE3:FVz4Sy9BZ~[rRg})05cks');
define('AUTH_SALT',        '(cG66RS6&@U7|5yj$]1WOf#TXh+8Q7@19<n[151G[/#x`Oq:!Y!}$3npFM#] v m');
define('SECURE_AUTH_SALT', '4cr4uyxWh_JxWgN;|qy~}t(3Rb69l d,;:/}ljoph@wqYS~z>-/ /C@,I2K #B1@');
define('LOGGED_IN_SALT',   '{BEn[]L*0BCjL%f|a.[t[+i1|R}Y4rnsVH`W,qfLzw>M39^d*y--,@uLR7#,n*k|');
define('NONCE_SALT',       'iW3+**XPHhueu!?tgv?}9}um8_)73$a.F:&B$p@|gz1rgyRUm#|04>kc=dd*.b,f');

/**#@-*/

/**
 * WordPress Database Table prefix.
 *
 * You can have multiple installations in one database if you give each a unique
 * prefix. Only numbers, letters, and underscores please!
 */
$table_prefix  = 'wp_';

/**
 * For developers: WordPress debugging mode.
 *
 * Change this to true to enable the display of notices during development.
 * It is strongly recommended that plugin and theme developers use WP_DEBUG
 * in their development environments.
 */
define('WP_DEBUG', false);

/* That's all, stop editing! Happy blogging. */

/** Absolute path to the WordPress directory. */
if ( !defined('ABSPATH') )
	define('ABSPATH', dirname(__FILE__) . '/');

/** Sets up WordPress vars and included files. */
require_once(ABSPATH . 'wp-settings.php');
