<?php
namespace PL2010\Laravel;

use PL2010\Laravel\Facades\MsGraph;
use PL2010\MsGraph\MsGraphManager;

use Illuminate\Contracts\Foundation\Application;
use Illuminate\Foundation\AliasLoader;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\ServiceProvider;
use Symfony\Component\Mailer\SentMessage;
use Symfony\Component\Mailer\Transport\AbstractTransport;

class MsGraphServiceProvider extends ServiceProvider {
	/** @var array Default naming scheme. */
	private array $naming = [
		'name' => 'msgraph',
		'alias' => 'MsGraph',
		'tk_repo' => 'oauth2_tokens',
		'mailer' => 'msgraph_mail',
	];

	public function register(): void {
		// Allow configuration to override naming scheme.
		$this->naming = array_merge(
			$this->naming,
			$this->app['config']['pl2010.msgraph.naming'] ?? []
		);

		// Make MsGraphManager available as 'msgraph' (or whatever
		// $this->naming['name'] specifies).
		$this->app->singleton($this->naming['name'], function() {
			$msgraph = new MsGraphManager([
				'error_log' => function(string $msg) {
					Log::error($msg);
				},
				'token_repo' => function() {
					return $this->app->make($this->naming['tk_repo']);
				},
			]);

			$name = $this->naming['name'];

			// Configure Graph clients, with first one as default.
			$first = true;
			foreach ($this->app['config'][$name] ?? [] as $client => $config) {
				// $config is either actual configuration array, or
				// a string that points to another config; no chaining
				if (is_string($config))
					$config = $this->app['config']["{$name}.{$config}"] ?? [];

				$msgraph->configure($client, $config, $first);
				$first = false;
			}

			return $msgraph;
		});

		// Make MsGraphManager available as facade 'MsGraph' (or whatever
		// $this->naming['alias'] specifies).
		AliasLoader::getInstance()->alias(
			$this->naming['alias'],
			MsGraph::class
		);
	}

	public function boot(): void {
		$name = $this->naming['name'];

		// Register MS Graph-based email transport 'msgraph'
		// (or whatever $this->naming['name'] specifies).
		$mailMgr = $this->app->make('mail.manager');
		$mailMgr->extend($name, function() {
			return new class(
				$this->naming['name'],
				$this->app
			) extends AbstractTransport {
				public function __construct(
					private string $name,
					private Application $app
				) {
					parent::__construct();
				}
				protected function doSend(SentMessage $message): void {
					$mailers = $this->app['config']['mail.mailers'] ?? null;
					$msgraph = $mailers[$this->name]['msgraph'] ?? null;
					$this->app
						->make($this->name)
						->get($msgraph)
						->sendEmail($message->getOriginalMessage());
				}
				public function __toString(): string {
					return $this->name;
				}
			};
		});
	}
}
