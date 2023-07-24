<?php
namespace PL2010\MsGraph;

use PL2010\MsGraph\Contracts\MsGraphDispatch;
use PL2010\OAuth2\Contracts\TokenRepository;

use Microsoft\Graph\Graph;
use Microsoft\Graph\Exception\GraphException;
use Microsoft\Graph\Http\GraphRequest;
use Symfony\Component\Mime\Message;
use Symfony\Component\Mime\Header\DateHeader;
use Symfony\Component\Mime\Header\HeaderInterface;
use Symfony\Component\Mime\Header\MailboxHeader;
use Symfony\Component\Mime\Header\MailboxListHeader;

use Closure;
use DateTime;
use DateTimeInterface;
use DateTimeZone;

/**
 * Implementation of the {@link MsGraphDispatch} interface.
 */
class MsGraphClient implements MsGraphDispatch {
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
		], $this->config);
	}

	/**
	 * Format date/time to Microsoft Graph API DateTimeOffset.
	 * @param DateTimeInterface $dt
	 * @return string UTC time string e.g. '2014-01-01T00:00:00Z'.
	 */
	private function formatDateTimeOffset(DateTimeInterface $dt): string {
		$utc = new DateTime('now', new DateTimeZone('UTC'));
		$utc->setTimestamp($dt->getTimestamp());
		return $utc->format('Y-m-d\\TH:i:sp');
	}

	/**
	 * Format message email header for MS Graph sendMail.
	 * @param Symfony\Component\Mime\Header\HeaderInterface $header
	 * @param bool $first Return only first address.
	 * @return array List of email addresses, or just the first one if $first.
	 */
	private function formatEmailAddresses(
		HeaderInterface $header,
		bool $first=false
	): array {
		// Extract adddresses from mailbox.
		$mboxAddresses = null;
		if ($header instanceof MailboxHeader) {
			$mboxAddresses[] = $header->getAddress();
		}
		else if ($header instanceof MailboxListHeader) {
			$mboxAddresses = $header->getAddresses();
		}

		$result = [];
		if ($mboxAddresses) {
			// Mailbox header addresses.
			foreach ($mboxAddresses as $address) {
				$result[] = [
					'emailAddress' => [
						'name' => $address->getName(),
						'address' => $address->getAddress(),
					],
				];
				if ($first) {
					$result = $result[0];
					break;
				}
			}
		}
		return $result;
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
	 * No support yet for multi-part messages.
	 */
	public function sendEmail(Message $msg, bool $save=false): void {
		$email = [];

		$body = $msg->getBody();
		$email['body'] = [
			'contentType' => strtolower($body->getMediaSubtype())==='html'
				? 'html'
				: 'text',
			'content' => $body->bodyToString(),
		];

		foreach ($msg->getHeaders()->all() as $header) {
			switch ($name = strtolower($header->getName())) {
			case 'date':
				assert($header instanceof DateHeader);
				$email['createdDateTime'] = $this->formatDateTimeOffset(
					$header->getDateTime()
				);
				break;
			case 'from':
			case 'sender':
				$email['from'] = $this->formatEmailAddresses($header, true);
				break;
			case 'bcc':
			case 'cc':
			case 'to':
				$email["{$name}Recipients"] =
					$this->formatEmailAddresses($header);
				break;
			case 'reply-to':
				$email['replyTo'] = $this->formatEmailAddresses($header);
				break;
			case 'subject':
				$email['subject'] = $header->getBodyAsString();
				break;
			default:
				break;
			}
		}

		$response = $this->prepareRequest(
			'POST',
			'/me/sendMail'
		)->attachBody([
			'message' => $email,
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
