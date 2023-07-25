<?php
namespace PL2010\MsGraph\Concerns;

use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\Part\AbstractPart;
use Symfony\Component\Mime\Part\AbstractMultipartPart;
use Symfony\Component\Mime\Part\DataPart;
use Symfony\Component\Mime\Part\TextPart;
use Symfony\Component\Mime\Part\Multipart\AlternativePart;

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
	 * @param string $textType Preferred text subtype among alternatives.
	 * @return array
	 */
	private function msGraphMessage(
		Email $email,
		string $textType='html'
	): array {
		$message = [];

		$body = $email->getBody();
		switch ($type = $body->getMediaType()) {
		case 'text':
			$message['body'] = $this->msGraphItemBody($body);
			break;
		case 'multipart':
			[
				$text,
				$attachments,
			] = $this->parseMultipart($body, $textType);
			$message['body'] = $this->msGraphItemBody($text);
			foreach ($attachments as $part) {
				$message['attachments'][] = $this->msGraphAttachment($part);
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

	/**
	 * Parse an alternative multi-part for the main text.
	 * @param \Symfony\Component\Mime\Part\Multipart\AlternativePart $multi
	 * @param string $textType Preferred subtype if not the first.
	 */
	private function parseAlternativeForText(
		AlternativePart $multi,
		string $textType=''
	): ?TextPart {
		$subtype = strtolower($textType);

		$first = null;
		$text = null;
		foreach ($multi->getParts() as $text) {
			// Non-text is unexpected, but let's continue.
			if (!$text instanceof TextPart)
				continue;
			$first ??= $text;

			if ($subtype === ''
				|| strtolower($text->getMediaSubtype()) === $subtype
			) {
				return $text;
			}
		}
		return $first;
	}

	/**
	 * Extract main message and attachments from multi-part.
	 * This expects the multi-part to be one of these forms:
	 * - alternative: TextPart alernatives
	 * - mixed: TextPart or AlternativePart followed by DataPart attachments
	 * @todo Handle related multipart.
	 * @param \Symfony\Component\Mime\Part\AbstractMultipartPart $multi
	 * @param string $textType Preferred text subtype among alternatives.
	 * @return array Tuple of text body and list of attachment parts.
	 */
	private function parseMultipart(
		AbstractMultipartPart $multi,
		string $textType='html'
	): array {
		$text = null;
		$attachments = [];

		$subtype = $multi->getMediaSubtype();
		$parts = $multi->getParts();
		switch ($subtype) {
		case 'alternative':
			// Main message in alternative form (html vs text).
			$text = $this->parseAlternativeForText($multi, $textType);
			break;
		case 'mixed':
			// Main message with attachments.
			// First is either a text or an alternatives of some.
			$first = reset($parts);
			if ($first instanceof TextPart) {
				$text = $first;
			}
			else if ($first instanceof AlternativePart) {
				$text = $this->parseAlternativeForText($first);
			}
			// The remaining parts are presumed to be attachments.
			$attachments = array_slice($parts, 1);
			break;
		default:
			break;
		}

		if (!$text) {
			throw new InvalidArgumentException(
				"no text part found in multipart/{$subtype}");
		}

		return [
			$text,
			$attachments,
		];
	}
}
