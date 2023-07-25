<?php
namespace PL2010\MsGraph\Concerns;

use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\Part\AbstractMultipartPart;

use DateTime;
use DateTimeImmutable;
use DateTimeZone;
use InvalidArgumentException;

/**
 * Build MS Graph Message.
 */
trait MsGraphMessageBuilder {
	/**
	 * Build MS Graph Attachment from Symfony MIME multi-part.
	 * @param \Symfony\Component\Mime\Part\AbstractMultipartPart $multi
	 * @return array
	 */
	private function msGraphAttachments(AbstractMultipartPart $multi): array {
		// TODO: large size attachment
		return [];
	}

	/**
	 * Build MS Graph DateTimeOffset from PHP date/time.
	 * @param \DateTimeImmutable $dt
	 * @return string UTC time string e.g. '2014-01-01T00:00:00Z'.
	 */
	private function msGraphDateTimeOffset(DateTimeImmutable $dt): string {
		$utc = new DateTime('now', new DateTimeZone('UTC'));
		$utc->setTimestamp($dt->getTimestamp());
		return $utc->format('Y-m-d\\TH:i:sp');
	}

	/**
	 * Build a MS Graph Message from Symfony MIME email.
	 * @param \Symfony\Component\Mime\Email $email
	 * @return array
	 */
	private function msGraphMessage(Email $email): array {
		$message = [];

		$body = $email->getBody();
		switch ($type = $body->getMediaType()) {
		case 'text':
			$message['body'] = [
				'contentType' => strtolower($body->getMediaSubtype()) === 'html'
					? 'html'
					: 'text',
				'content' => $body->bodyToString(),
			];
			break;
		case 'multipart':
			$message['attachments'] = $this->msGraphAttachments($body);
			break;
		default:
			throw new InvalidArgumentException(
				"unsupported email media type: $type");
		}

		if ($from = $email->getFrom()) {
			$message['from'] = $this->msGraphRecipient(reset($from));
		}
		if ($sender = $email->getSender()) {
			$message['sender'] = $this->msGraphRecipient($sender);
		}
		if ($date = $email->getDate()) {
			$message['createdDateTime'] = $this->msGraphDateTimeOffset($date);
		}
		if ($to = $email->getTo()) {
			foreach ($to as $addr)
				$message['toRecipients'][] = $this->msGraphRecipient($addr);
		}
		if ($cc = $email->getCc()) {
			foreach ($cc as $addr)
				$message['ccRecipients'][] = $this->msGraphRecipient($addr);
		}
		if ($bcc = $email->getBcc()) {
			foreach ($bcc as $addr)
				$message['bccRecipients'][] = $this->msGraphRecipient($addr);
		}
		if ($reply = $email->getReplyTo()) {
			foreach ($reply as $addr)
				$message['replyTo'][] = $this->msGraphRecipient($addr);
		}
		if (is_string($subject = $email->getSubject())) {
			$message['subject'] = $subject;
		}

		return $message;
	}

	/**
	 * Build MS Graph Recipient from Symfony MIME address.
	 * @param \Symfony\Component\Mime\Address $addr
	 * @return array
	 */
	private function msGraphRecipient(Address $addr): array {
		return [
			'emailAddress' => [
				'name' => $addr->getName(),
				'address' => $addr->getAddress(),
			],
		];
	}
}
