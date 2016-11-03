
const WHITELIST_LABEL = 'Shared Excluded Placements';
const SPREADSHEET_ID = '1IVRydmnDQ3WLVKk5McRf45-YwY2U6Gr8ucnUCVix4Fw';
const SHARED_PLACEMENT_LIST_NAME = 'Shared Excluded Placements';


function run () {
  // retrieve placements to exclude
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const placements = sheet.getDataRange().getValues().slice(1).map<string>(value => value[0] as string);
  //Logger.log(placements);

  // retrieve shared excluded placement list
  const excludedPlacementListIter = AdWordsApp.excludedPlacementLists()
    .withCondition(`Name = "${SHARED_PLACEMENT_LIST_NAME}"`)
    .withLimit(1)
    .get();

  let sharedExcludedPlacementList: AdWordsScripts.AdWordsApp.ExcludedPlacementList = undefined;
  if (excludedPlacementListIter.hasNext()) {
    sharedExcludedPlacementList = excludedPlacementListIter.next();

    // update the shared excluded placement list
    sharedExcludedPlacementList.addExcludedPlacements(placements);

    // retrieve campaigns that are not attached to the shared excluded placement list
    const campaignsSharedToExcludedPlacementList = [];
    const sharedCampaignIter = sharedExcludedPlacementList.campaigns()
      .withCondition("Status = ENABLED").get();
    while (sharedCampaignIter.hasNext()) {
      const campaign = sharedCampaignIter.next();
      campaignsSharedToExcludedPlacementList.push(campaign.getId());
    }
    // share excluded placement list to accounts not already shared
    const campaignIter = AdWordsApp.campaigns()
      .withCondition("Status = ENABLED")
      .withCondition(`Id NOT_IN [${campaignsSharedToExcludedPlacementList.map(v => `"${v}"`).join(',')}]`)
      .get();
    while (campaignIter.hasNext()) {
      const campaign = campaignIter.next();
      campaign.addExcludedPlacementList(sharedExcludedPlacementList);
    }
  } else {
    // create a new shared placement list
    const sharedExcludedPlacementListOperation = AdWordsApp.newExcludedPlacementListBuilder()
      .withName(SHARED_PLACEMENT_LIST_NAME).build();
    sharedExcludedPlacementList = sharedExcludedPlacementListOperation.getResult();

    // update the shared excluded placement list
    sharedExcludedPlacementList.addExcludedPlacements(placements);

    // share excluded placement list to accounts not already shared
    const campaignIter = AdWordsApp.campaigns().withCondition("Status = ENABLED").get();
    while (campaignIter.hasNext()) {
      const campaign = campaignIter.next();
      campaign.addExcludedPlacementList(sharedExcludedPlacementList);
    }
  }
}


function executeInSequence (sequentialIds, executeSequentiallyFunc) {
  Logger.log('Executing in sequence : ' + sequentialIds);
  sequentialIds.forEach(function (accountId) {
    const account = MccApp.accounts().withIds([accountId]).get().next();
    MccApp.select(account);
    executeSequentiallyFunc();
  });
}


function main () {
  try {
    const accountIterator = MccApp.accounts()
      .withCondition(`LabelNames CONTAINS '${WHITELIST_LABEL}'`)
      .orderBy('Name')
      .get();
    // map account entities to
    const accountIds = [];
    while (accountIterator.hasNext()) {
      const account = accountIterator.next();
      accountIds.push(account.getCustomerId());
    }
    const parallelIds = accountIds.slice(0, 50);
    const sequentialIds = accountIds.slice(50);
    // execute accross accounts
    MccApp.accounts()
      .withIds(parallelIds)
      .executeInParallel('run');
    if (sequentialIds.length > 0) {
      executeInSequence(sequentialIds, run);
    }
  } catch (exception) {
    // not an Mcc
    Logger.log('Running on non-MCC account.');
    run();
  }
}
