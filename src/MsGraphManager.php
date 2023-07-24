<?php
namespace PL2010\MsGraph;

use PL2010\MsGraph\MsGraphClient;
use PL2010\MsGraph\Contracts\MsGraphDispatch;

use Symfony\Component\Mime\Message;

/**
 * Manager of MS Graph API clients.
 * Clients are retreived by name using the {@link get()} method.
 * This class also is facade for the default provider.
 */
class MsGraphManager implements MsGraphDispatch {
	/** @var array MS Graph dispatch config or actual client by name. */
	protected array $dispatches = [];

	/** @pvar ?string Default dispatch name. */
	protected ?string $default = null;

	public function __construct(
		/**
		 * Options including callback 'error_log'.
		 * The options are passed to {@link MsGraphClient}.
		 **/
		protected array $options=[]
	) {
	}

	/**
	 * Configure a MS Graph dispatch.
	 */
	public function configure(
		string $name,
		array $config,
		bool $default=false
	): static {
		$this->dispatches[$name] = $config;
		if ($default)
			$this->default = $name;
		return $this;
	}

	/**
	 * Get MS Graph dispatch by name.
	 */
	public function get(?string $name=null): ?MsGraphDispatch {
		$name = $name
			?? $this->default
			?? array_keys($this->dispatches)[0] ?? null;
		$impl = $this->dispatches[$name] ?? null;
		if ($impl instanceof MsGraphDispatch || $impl === null)
			return $impl;

		$impl = new MsGraphClient($name, $impl, $this->options);
		$this->dispatches[$name] = $impl;
		return $impl;
	}

	public function sendEmail(Message $msg, bool $save=false): void {
		$this->get()->sendEmail($msg, $save);
	}
}
