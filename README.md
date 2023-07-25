# Microsoft Graph Utilities for PHP #

[packagist]: https://packagist.org/
[plai2010-oauth2-client]: https://packagist.org/packages/plai2010/php-oauth2

This is a utility package for using Microsoft Graph API in PHP
applications. In this version it only supports sending email
messages to the `sendMail` endpoint.

The use case for this package was email notification in a Laravel
(10.x) application. A Laravel service provider is included.

This package is designed to work with
[plai2010/php-oauth2][plai2010-oauth2] for OAuth2 token management.

## Installation ##

The package can be installed from [Packagist][packagist]:
```
$ composer require plai2010/php-msgraph
```

One may also clone the source repository from `Github`:
```
$ git clone https://github.com/plai2010/php-msgraph.git
$ cd php-msgraph
$ composer install
```

## Example: Sending Email ##

This package works best as a Laravel extension. It is _possible_ to
use it as just a utility library, but as one would see in this example,
it can be a little awkard.
```
use PL2010\MsGraph\MsGraphClient;
use PL2010\OAuth2\Contracts\TokenRepository;

// Some PL2010\OAuth2\Contracts\TokenRepository.
// Here we have a single hard coded access token
// wrapped as a token repository.
$tkrepo = new class implements TokenRepository {
	public function getOAuth2Token(string $key, int $valid=0): array {
		return [
			// One single hard coded token regardless of $key.
			'access_token' => '....',
		];
	}
	public function putOAuth2Token(string $key, array $token): static {
		return $this;
	}
};

$msgraph = new MsGraphClient('sendmail', [
	// OAuth2 access token will be retrieved from the
	// token repository using this key.
	'token_key' => 'does-not-matter',
], [
	// Provide token repository.
	'token_repo' => $tkrepo,
]);

$msg = (new \Symfony\Component\Mime\Email())
	->from('no-reply@example.com')
	->to('john.doe@example.com')
	->subject('hello')
	->html('<h3>Greetings!</h3>')
	->attach(file_get_contents('wave.png'), 'wave.png', 'image/png')
;
$msgraph->sendEmail($msg);
```

## Client Manager ##

An application typically acts as a single MS Graph API client.
[`MsGraphManager`](src/OMsGraphManager.php) allows multiple
clients to be configured for more complex scenarios. For example,
there may a client that acts for the application, and another
as agent for individual users.
```
// Create manager.
$manager = new PL2010\MsGraph\MsGraphManager;

// Configure a application/service level client.
$manager->configure('service', [
	'token_key' => 'msgraph:app',
]);

// Configure an agent client. Here we assume that
// the token repository recognizes the '@user' tag
// in the token key and looks up user specific
// access token.
$manager->configure('agent', [
	'token_key' => 'msgraph:agent@user',
]);
```

## Laravel Integration ##

This package includes a Laravel service provider that makes
a singleton [`MsGraphManager`](src/MsGraphManager.php) available
two ways:

	* Abstract 'msgraph' in the application container, i.e. `app('msgraph')`.
	* Facade alias 'MsGraph', i.e. `MsGraph::`.

The configuration file is `config/msgraph.php`. It returns an
associative array of client configurations by names, like this:
```
<?php
return [
	'msgraph_site' => [
		'token_key' => 'msgraph:app',
	],
	'msgraph_ua' => [
		'token_key' => 'msgraph:agent@user',
	],
	...
];
```

Also a 'msgraph' email transport is registered. So one may define a
mailer in `config/mailer.php` like this:
```
	'outlook' => [
		'transport' => 'msgraph',

		// Points to entry in config/msgraph.php.
		'msgraph' => 'msgraph_site',
	]
```

A message can then be sent like this:
```
	use Illuminate\Mail\Mailable;
	$msg = (new Mailable())
		->to('john.doe@example.com')
		->subject('hello')
		->html('<h3>Greetings!</h3>')
		->send(Mail::mailer('outlook'));
```
