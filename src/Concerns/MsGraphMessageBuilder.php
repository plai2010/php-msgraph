<?php
namespace PL2010\MsGraph\Concerns;

use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\Part\AbstractPart;
use Symfony\Component\Mime\Part\DataPart;
use Symfony\Component\Mime\Part\TextPart;

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
	 * @todo Handle large size attachment (over 3MB).
	 * @param \Symfony\Component\Mime\Part\AbstractPart $part
	 * @return array
	 */
	private function msGraphAttachment(AbstractPart $part): array {
		$data = $part instanceof TextPart
			? $part->getBody()
			: $part->bodyToString();
		$attachment = [
			'@odata.type' => '#microsoft.graph.fileAttachment',
			'contentBytes' => base64_encode($data),
			'contentType' => $part->getMediaType().'/'.$part->getMediaSubtype(),
			'size' => strlen($data),
		];
		if ($part instanceof TextPart) {
			if (($name = $part->getName()) !== null)
				$attachment['name'] = $name;
			if ($part->getDisposition() === 'inline')
				$attachment['isInline'] = true;
		}
		if ($part instanceof DataPart) {
			if ($part->hasContentId())
				$attachment['contentId'] = $part->getContentId();
			if (($filename = $part->getFilename()) !== null)
				$attachment['name'] = $filename;
		}
		return $attachment;
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
	 * Build MS Graph ItemBody from Symfony MIME part.
	 * @param \Symfony\Components\Mime\Part\TextPart $part
	 * @return array
	 */
	private function msGraphItemBody(TextPart $part): array {
		return [
			'contentType' => strtolower($part->getMediaSubtype()) === 'html'
				? 'html'
				: 'text',
			'content' => $part->getBody(),
		];
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
			$message['body'] = $this->msGraphItemBody($body);
			break;
		case 'multipart':
			$text = null;
			foreach ($body->getParts() as $part) {
				if (!$text && $part instanceof TextPart) {
					// First text part is main body.
					$text = $part;
					$message['body'] = $this->msGraphItemBody($part);
				}
				else {
					$message['attachments'][] = $this->msGraphAttachment($part);
				}
			}
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
