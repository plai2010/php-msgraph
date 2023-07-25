<?php
namespace PL2010\Laravel\Facades;

use Illuminate\Support\Facades\Facade;

/**
 * Facade for MS Graph API invocation.
 * @method static void sendEmail(\Symfony\Component\Mime\Email $email, bool $save=false);
 */
class MsGraph extends Facade {
	/**
	 * Get the registered name of the component.
	 */
	protected static function getFacadeAccessor(): string {
		return config('pl2010.msgraph.naming.name', 'msgraph');
	}
}
