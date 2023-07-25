<?php
namespace PL2010\MsGraph;

use PL2010\MsGraph\Concerns\MsGraphMessageBuilder;
use PL2010\MsGraph\Contracts\MsGraphDispatch;
use PL2010\OAuth2\Contracts\TokenRepository;

use Microsoft\Graph\Graph;
use Microsoft\Graph\Exception\GraphException;
use Microsoft\Graph\Http\GraphRequest;
use Symfony\Component\Mime\Email;
use Closure;

/**
 * Implementation of the {@link MsGraphDispatch} interface.
 */
class MsGraphClient implements MsGraphDispatch {
	use MsGraphMessageBuilder;

	/** @var TokenRepository|Closure Token repository or resolver. */
	protected TokenRepository|Closure $tk_repo;

	public function __construct(
		private string $name,
		private array $config,
		array $options=[]
	) {
		if (is_callable($logger = $options['error_log'] ?? null))
			$this->err_log = $logger;
		$this->tk_repo = $options['token_repo'] ?? function() {
			$factory = $this->config['token_repo_factory'] ?? null;
			if ($factory && is_callable($factory)) {
				$key = $this->config['token_key'];
				return $factory($key);
			}
			return null;
		};

		$this->config = array_merge([
			'token_key' => $this->name,
			'token_ttl' => 120,
			'text_subtype' => 'html',
		], $this->config);
	}

	/**
	 * Get access token from token repository for this client.
	 * @return ?string
	 */
	private function getAccessToken(): ?string {
		$tkrepo = $this->getTokenRepository();
		if (!$tkrepo)
			return null;

		$key = $this->config['token_key'];
		$ttl = $this->config['token_ttl'];
		$token = $tkrepo->getOAuth2Token($key, $ttl);
		return $token['access_token'] ?? null;
	}

	/**
	 * Get token repository for this client.
	 * @return ?\PL2010\OAuth2\Contracts\TokenRepository
	 */
	private function getTokenRepository(): ?TokenRepository {
		// Token repository instance already?
		$tkrepo = $this->tk_repo;
		if ($tkrepo instanceof TokenRepository)
			return $tkrepo;

		// Attempt to resolve key to repository.
		$tkrepo = ($tkrepo)($this->config['token_key']);
		if ($tkrepo instanceof TokenRepository)
			return $this->tk_repo = $tkrepo;

		$this->logError("no token repository for MSGraphClient: {$this->name}");
		return null;
	}

	/**
	 * Prepare a new MS Graph request.
	 * The request will have access token set already.
	 */
	private function prepareRequest(
		string $method,
		string $endpoint
	): GraphRequest {
		$request = (new Graph())
			->setAccessToken($this->getAccessToken())
			->createRequest($method, $endpoint);
		if ($timeout = $this->config['timeout'] ?? null)
			$request->setTimeout($timeout);
		return $request;
	}

	/**
	 * {@inheritdoc}
	 */
	public function sendEmail(Email $email, bool $save=false): void {
		$message = $this->msGraphMessage($email, $this->config['text_subtype']);

		$response = $this->prepareRequest(
			'POST',
			'/me/sendMail'
		)->attachBody([
			'message' => $message,
			'saveToSentItems' => $save,
		])->execute();

		$status = $response->getStatus();
		if ($status !== 202) {
			$this->logError("failed to send email with MS Graph: status=$status");
			throw new GraphException('failed to send email');
		}
	}

	/** @var callable Error logger. */
	protected $err_log = 'error_log';

	protected function logError(string $msg): void {
		($this->err_log)($msg);
	}
}
