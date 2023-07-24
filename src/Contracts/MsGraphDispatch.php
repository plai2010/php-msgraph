<?php
namespace PL2010\MsGraph\Contracts;

use Symfony\Component\Mime\Message;

/**
 * MS Graph API invocation.
 */
interface MsGraphDispatch {
	/**
	 * Send an email message.
	 * @param bool $save Save message to "Sent Items" after sent.
	 */
	public function sendEmail(Message $msg, bool $save=false): void;
}
