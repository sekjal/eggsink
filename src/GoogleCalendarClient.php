<?php

class GoogleCalendarClient
{
    private $client = null;

    private $scopes = [
        'https://www.googleapis.com/auth/calendar'
    ];

    private $calendarId = null;

    public function __construct($clientId, $email, $keyFile, $appName, $calendarId=null)
    {
        $client = new Google_Client();
        $client->setApplicationName($appName);
	$client->setAuthConfig($keyFile);
	$client->setScopes($this->scopes);
/*
	$scope = new Google_Service_Calendar_AclRuleScope();
	$scope->setType('user');
	$scope->setValue( $email );

	$rule = new Google_Service_Calendar_AclRule();
	$rule->setRole( 'owner' );
	$rule->setScope( $scope );
*/
        $this->client = new Google_Service_Calendar($client);
//	$result = $this->client->acl->insert($calendarId, $rule);

        $this->calendarId = $calendarId;


    }

    /**
     * @param int $daysFromNow
     * @return array
     */
    public function getEvents($daysFromNow)
    {
        $options = [
            'singleEvents' => true,
            'orderBy' => 'startTime',
            'timeMin' => date(DATE_RFC3339),
            'timeMax' => date(DATE_RFC3339, strtotime('+' . $daysFromNow . ' days'))
        ];
        $events = $this->client->events->listEvents($this->calendarId, $options);
        $items = [];
        /** @var Google_Service_Calendar_Event $event */
        foreach ($events as $i => $event) {
            if (!empty($event->getStart()['date'])) {
                $start = date('Y-m-d', strtotime($event->getStart()['date']));
                $end = date('Y-m-d', strtotime($event->getEnd()['date']));
                $isAllDayEvent = true;
            } else {
                $start = $event->getStart()['dateTime'];
                $end = $event->getEnd()['dateTime'];
                $isAllDayEvent = false;
            }

            if (!empty($event->getExtendedProperties()['private']['ewsId'])) {
                $private = $event->getExtendedProperties()['private'];
                $ewsId = $private['ewsId'];
                $ewsChangeKey = $private['ewsChangeKey'];
            } else {
                $ewsId = null;
                $ewsChangeKey = null;
            }

            $items[] = [
                'id' => $event->getId(),
		'etag' => $event->getEtag(),
                'subject' => $event->getSummary(),
                'start' => $start,
                'end' => $end,
                'isAllDayEvent' => $isAllDayEvent,
		'isBusyStatus' => ($event->getTransparency == 'opaque'),
                'location' => $event->getLocation() ? $event->getLocation() : null,
                'ewsId' => $ewsId,
                'ewsChangeKey' => $ewsChangeKey
            ];
        }

        return $items;
    }

    public function addEvent(array $details = [])
    {
        $event = $this->buildEvent($details);

        return $this->client->events->insert($this->calendarId, $event);
    }

    public function updateEvent($eventId, array $details = [])
    {
        $event = $this->buildEvent($details);

        return $this->client->events->update($this->calendarId, $eventId, $event);
    }

    public function deleteEvent($eventId)
    {
        return $this->client->events->delete($this->calendarId, $eventId);
    }

    private function buildEvent(array $details = [])
    {
        $event = new Google_Service_Calendar_Event();
        $event->setSummary($details['subject']);
        $event->setLocation($details['location']);

        $start = new Google_Service_Calendar_EventDateTime();
        $end = new Google_Service_Calendar_EventDateTime();
        if ($details['isAllDayEvent']) {
            $start->setDate(date('Y-m-d', strtotime($details['start'])));
            $end->setDate(date('Y-m-d', strtotime($details['end'])));
        } else {
            $start->setDateTime($details['start']);
            $end->setDateTime($details['end']);
        }
        $event->setStart($start);
        $event->setEnd($end);

        $extendedProperties = new Google_Service_Calendar_EventExtendedProperties();
        $extendedProperties->setPrivate([
            'ewsId' => $details['ewsId'],
            'ewsChangeKey' => $details['ewsChangeKey']
        ]);
        $event->setExtendedProperties($extendedProperties);

        return $event;
    }

}


