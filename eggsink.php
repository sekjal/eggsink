<?php

require_once dirname(__FILE__) . '/autoload.php';
require_once dirname(__FILE__) . '/config/config.php';

addToLog('Started eggsink');

// Step 1: Connect to Serves and Fetch Events

// Step 1a:  Connect to Exchange
for ($i=1; $i<=RETRIES; $i++) {
    try
    {
        $exchange = new ExchangeClient(EXCHANGE_SERVER, EXCHANGE_USERNAME, EXCHANGE_PASSWORD);
        $exchangeevents = $exchange->getCalendarEvents(SYNC_DAYS_FROM_NOW);
        break;
    } catch (Exception $e) {
        if ($i==$retries)
        {
            addToLog('Failed to connect to Exchange. Please try again later.');
	    addToLog($e);
            exit(0);
        }
    }
    addToLog('Exchange connection retry ' . $i);
}

// Step 1b:  Connect to Google Calendar
$google = new GoogleCalendarClient(GOOGLE_CLIENT_ID, GOOGLE_EMAIL, dirname(__FILE__) . '/config/' . GOOGLE_KEY_FILE, 'EggSink', GOOGLE_CALENDAR_ID);
$googleevents = $google->getEvents(SYNC_DAYS_FROM_NOW);


// Step 2: Divide Native Events from Imported Events

// Step 2a: Split Google events to those that were previously exported from Exchange
// and those created native on Google Calendar
$exchangeEventsOnGoogle = [];
$googleNativeEvents = [];
foreach ($googleevents as $event) {
    if ( empty($event['ewsId']) ) {
	// if there is no ewsId, then it's Google Native
	// key on googlecalendarid
        $googleNativeEvents[$event['id']] = $event;
    } else {
        // otherwise, assume it's from Exchange
	// key on ewsId
	$exchangeEventsOnGoogle[$event['ewsId']] = $event;
    }
}

// Step 2b: Split Exchange events to those that were previously exported from Google
// and those created native on Exchange
$googleEventsOnExchange = [];
$exchangeNativeEvents = [];
foreach ($exchangeevents as $event) {
    if ( empty($event['googlecalendarid']) ) {
        $exchangeNativeEvents[$event['id']] = $event;
    } else {
	$googleEventsOnExchange[$event['googlecalendarid']] = $event;
    }
}

// Step 3: Add update events

// Step 3a: Add or update events from Exchange to Google
foreach ($exchangeNativeEvents as $meeting) {
    if (!$meeting['isBusyStatus'] || !$meeting['isPublic']) {
        continue;
    }

    $details = [
        'subject' => $meeting['subject'],
        'location' => $meeting['location'],
        'start' => $meeting['start'],
        'end' => $meeting['end'],
        'isAllDayEvent' => $meeting['isAllDayEvent'],
        'ewsId' => $meeting['id'],
        'ewsChangeKey' => $meeting['changeKey']
    ];


    if (!empty($exchangeEventsOnGoogle[$meeting['id']])) { // This an existing meeting

        if ($exchangeEventsOnGoogle[$meeting['id']]['ewsChangeKey'] != $meeting['changeKey']) {
            $google->updateEvent($exchangeEventsOnGoogle[$meeting['id']]['id'], $details);
        }
        unset($exchangeEventsOnGoogle[$meeting['id']]); // remove the event from the remaining events queue
	addToLog("Exchange event '" . $details['subject'] . "' updated on Google");
    } else { // This must be a new meeting

        $event = $google->addEvent($details);
	addToLog("Exchange event '" . $details['subject'] . "' created on Google");
    }
}
// Delete remaining Exchnage events from Google
foreach ($exchangeEventsOnGoogle as $event) {
    $google->deleteEvent($event['id']);
    addToLog("Exchange event '" . $event['subject'] . "' deleted from Google");
}


// Step 3b: Add or update events from Exchange to Google
foreach ($googleNativeEvents as $meeting) {
    $details = [
	'googlecalendarid' => $meeting['id'],
        'subject' => $meeting['subject'],
        'location' => $meeting['location'],
        'start' => $meeting['start'],
        'end' => $meeting['end'],
    ];

    if (!empty($googleEventsOnExchange[$meeting['id']])) { // This an existing meeting
        $exchange->updateEvent($googleEventsOnExchange[$meeting['id']]['id'], $details);
        unset($googleEventsOnExchange[$meeting['id']]); // remove the event from the remaining events queue
	addToLog("Google event '" . $details['subject'] . "' updated on Exchange");
    } else { // This must be a new meeting

        $event = $exchange->addEvent($details);
	addToLog("Google event '" . $details['subject'] . "' created on Exchange");
    }

}

// Delete remaining Exchnage events from Google
foreach ($googleEventsOnExchange as $event) {
    $exchange->deleteEvent($event['id'], $event['changeKey']);
    addToLog("Google event '" . $event['subject'] . "' deleted from Exchange");
}




addToLog('Finished eggsink');

function addToLog($message)
{
    echo '[' . date(DATE_RFC3339) . '] ' . $message . "\n";
}
