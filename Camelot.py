import pdb
import copy
import os
import random
import shutil
import sys

from docx import Document
from docx.shared import Pt, RGBColor, Length, Inches
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def insertHR(paragraph):
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)

# get_role_descriptions - this is called when information files are generated.
def get_role_description(role):
    return {
        'Tristan' : 'The person you see is also Good and is aware that you are Good.\nYou and Iseult are each a valid Assassination target. You and Iseult each know one location that Excalibur is not.',
        'Iseult' : 'The person you see is also Good and is aware that you are Good.\nYou and Tristan are each a valid Assassination target. You and Tristan each know one location that Excalibur is not.',
        'Merlin' : 'You know which people have Evil roles, but not who has any specific role.\n\nYou are a valid Assassination target.\n\nWhile Excalibur is in the Lake, you may play Reversal cards on quests.',
        'Percival' : 'You know which people have the Merlin and Morgana roles, but not who has each.\n\nWhile Excalibur is at Camelot, you may play Reversal cards on quests.',
        'Arthur' : 'You must obtain Excalibur to ensure victory for your forces. To do this, you must search for Excalibur in one or more or three locations: The Stone, The Lake, or At Camelot.\nYou may declare as a Sword Seeker, Arthur, or Accolon (your choice) to begin your search. After you declare, each time you are given the Scabbard, you may search the current location for Excalibur instead of changing the location. If it is in that location, you gain Excalibur.\nOnce you have Excalibur, when you go on a quest, you may expend its power to force a quest to succeed even when it would otherwise fail. You must expend Excalibur before the cards are read, but after they have been given to the Leader.\nOnce you expend Excalibur\'s power, you can no longer use this ability.',
        'Titania' : 'You appear as Evil to all players with Evil roles (except Colgrevance).',
        'Nimue' : 'You are a member of the Fairy Court. You can only play Regrowth cards on quests. If you are given The Holy Grail, you must play Success cards unless Excalibur is in the Lake, in which case, you may still play Regrowth cards.\n\nIf The Holy Grail is ever fully corrupted, you win the game unless the mortals can identify you and the other members of the Fairy Court. After The Holy Grail is corrupted, you may play any card you want on any quests you attend.',
        'Galahad' : 'You can gain the Reveal power only by going on the First or Fourth Quest.\n\nREVEAL:\nWhile Excalibur is at Camelot, you may declare as a Knight of Camelot, Galahad, or Lancelot (your choice). When you do, you instruct all players to close their eyes and hold their fists out in front of them. Name one good role. If a player has that role, they must raise their thumb to indicate they are playing that role. You can then instruct all players to put their hands down, open their eyes, and resume play normally.',
        'Guinevere' : 'You know two \"rumors\" about other players, but (with the exception of Arthur) nothing about their roles.\n\nThese rumors give you a glimpse at somebody else\'s character information, telling you who they know something about, but not what roles they are.\n\nFor instance, you if you heard a rumor about Player A seeing Player B, it might mean Player A is Merlin seeing an Evil player, or it might mean they are both Evil and can see each other.',
        'Lamorak' : 'You can see two pairs of players.\nOne pair of players are against each other (Good and Evil or Pelinor and the Questing Beast), and the other pair are on the same side (Evil and Evil or Good and Good).',
        'Ector' : 'You know which Good roles are in the game, but not who has any given role.',
        'Dagonet' : 'You cannot speak, but can communicate through gibberish sounds and body language.\n\nYou know Arthur.\n\nYou appear Evil to Merlin and to all Evil players.\n\nOnly Ector may know if Dagonet is in this game.',
        'Uther' : 'You can gain the Exile power by either voting against your own quest proposal while you are the leader, or by voting against a quest proposal you have been chosen to attend.\n\nEXILE:\nWhile Excalibur is in the Stone, you may declare as a King of the Realm, Uther, or Vortigurn (your choice). You may only do this after a new leader is selected but before a quest vote occurs. If you do, you may select one player to be exiled from the game until the next quest is completed. That player is required to view your role information, and will see which role you possess. The exiled player must leave the play area to view this information, and you must be the one to go retrieve that player, affording you a moment of privacy with that player, if you wish.',
        'Bedivere' : 'You can gain the Suspend power by choosing not to move Excalibur when you have the Scabbard.\n\nSUSPEND:\nWhile Excalibur is in the Lake, after the quest cards have been collected for a quest, but before they are read, you can declare as a Guardian of Truth, Bedivere, or Agravaine (your choice). If you do, you may look at the quest cards before the leader and remove one of them. The next time you attend a quest, you MUST play that card.',
        'Gawain' : 'You know all members of the Fairy Court, Good and Evil. Your presence has caused the Grail to start slightly corrupted.\n\nYou are a valid assassination target.\n\nEXTRA:\nIf there are NO members of the Fairy Court in this game, you may play Reversal cards on quests.',
        'Bors' : 'You may play Cleanse cards on quests. The Cleanse cards do not count as Successes or Failures, but remove any Regrowth or Rot cards from the quest cards. If Cleanse is the only remaining card, it counts as a Failure. If Cleanse does not remove any Regrowth or Rot cards, it counts as a Failure. Keep a secret tally of how many Regrowth or Rot cards you removed. When that number reaches three or higher, you may declare as Bors to claim the Holy Grail. If you do this, the Final Quest can only Fail if two or more Failure cards are played.',
        'Bertilak' : 'You are a member of the Fairy Court. You can only play Regrowth cards on quests. If you are given The Holy Grail, you must play Rot cards instead.\n\nIf The Holy Grail is ever fully corrupted, you win the game unless the mortals can identify you and the other members of the Fairy Court. After The Holy Grail is corrupted, you may play any card you want on any quests you attend.',

        'Mordred' : 'You are hidden from all Good roles that could reveal that information.\nLike other Evil characters, you know who else is Evil (except Colgrevance).',
        'Morgana' : 'You appear like Merlin to Percival.\n\nWhile Excalibur is in the Stone, you may play Reversal cards on quests.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).',
        'Maleagant' : 'You must play a Fail card on each mission you attend.\n\nEach time you are given the Scabbard, you may declare as Maleagant to force the next quest to have one additional knight attend it. This ability cannot be used on the final quest. This ability cannot be used if you use the Leader role to take the Scabbard.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).',
        'Agravaine' : 'You have the Suspend power.\n\nSUSPEND:\nWhile Excalibur is in the Lake, after the quest cards have been collected for a quest, but before they are read, you can declare as a Guardian of Truth, Bedivere, or Agravaine (your choice). If you do, you may look at the quest cards before the leader and remove one of them. The next time you attend a quest, you MUST play that card.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).',
        'Colgrevance' : 'You know not only who else is Evil, but what role each other Evil player possesses.\nEvil players know that there is a Colgrevance, but do not know that it is you or even that you are Evil.',
        'Accolon' : 'You know Arthur and must beat him to Excalibur. Arthur does not know Excalibur\'s location, but does know which players gain power by having Excalibur there.\n\nTo claim Excalibur, you must gain possession of the Scabbard twice by any means. Once you do, you may search the current location for Excalibur rather than moving it. If it is in that location, you gain Excalibur immediately. On any future quests you attend, you may expend Excalibur\'s power to cause that Quest to fail, even when it would succeed.\n\nIf you wish, you may choose to declare as a Sword Seeker, Arthur, or Accolon (your choice), though you are not required to do so.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).',
        'Lancelot' : 'You win if either three missions fail OR two missions fail and the assassination attempt at the end of the game fails. If only one quest fails or no quests fail, you are a valid assassination target. You cannot win if an assassination attempt succeeds.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).\n\nYou have the Reveal power.\n\nREVEAL:\nWhile Excalibur is at Camelot, you may declare as a Knight of Camelot, Galahad, or Lancelot (your choice). When you do, you instruct all players to close their eyes and hold their fists out in front of them. Name one good role. If a player has that role, they must raise their thumb to indicate they are playing that role. You can then instruct all players to put their hands down, open their eyes, and resume play normally.',
        'Vortigurn' : 'You have the Exile power.\n\nEXILE:\nWhile Excalibur is in the Stone, you may declare as a King of the Realm, Uther, or Vortigurn (your choice). You may only do this after a new leader is selected but before a quest vote occurs. If you do, you may select one player to be exiled from the game until the next quest is completed. That player is required to view your role information, and will see which role you possess. The exiled player must leave the play area to view this information, and you must be the one to go retrieve that player, affording you a moment of privacy with that player, if you wish.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).',
        'Annowre' : 'You know where Excalibur may be retrieved. You lose the game if anyone expends Excalibur\'s power.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).',
        'Palamedes' : 'You know Iseult and you have the Exile power.\nIf you Exile Tristan, you can no longer lose the game unless he or Iseult are assassinated. Tristan will then be required to play only Failures on any quests he attends, and you will only be allowed to play Success cards on quests. Tristan and Iseult remain assassination targets and if either of them are assassinated, you lose the game.\n\nEXILE:\nWhile Excalibur is in the Stone, you may declare as a King of the Realm, Uther, or Vortigurn (your choice). You may only do this after a new leader is selected but before a quest vote occurs. If you do, you may select one player to be exiled from the game until the next quest is completed. That player is required to view your role information, and will see which role you possess. The exiled player must leave the play area to view this information, and you must be the one to go retrieve that player, affording you a moment of privacy with that player, if you wish.\n\nLike other Evil characters, you know who else is Evil (except Colgrevance).\n\nEXILED PLAYER:\nIf you are Tristan, Palamedes is challenging you for the hand of your lady love. You are still Good but can only play Failure cards on quests. Palamedes is still Evil but can only play Success cards on quests. You and Iseult are still valid asssassination targets, so you must keep your role secret if you hope to win.',
        'Mab' : 'You are a member of the Fairy Court. You can only play Rot cards on quests. You ignore the effects of the Holy Grail.\n\nIf The Holy Grail is ever fully corrupted, you win the game unless the mortals can identify you and the other members of the Fairy Court. After The Holy Grail is corrupted, you may play any card you want on any quests you attend.',
        'Oberon' : 'You are a member of the Fairy Court. You can play Rot or Regrowth cards on quests. You know Gawain and Bertilak.\n\nIf The Holy Grail is ever fully corrupted, you win the game unless the mortals can identify you and the other members of the Fairy Court. After The Holy Grail is corrupted, you may play any card you want on any quests you attend.',

        'Pelinor' : 'You are Neutral in this battle and have no allies in this game.\n\nYour nemesis is The Questing Beast, who is also Neutral.\n\nCARDS YOU CAN PLAY:\n> \"Success\"\n> \"Reversal\"\n\nTO WIN:\n> The Fifth Quest must occur and you must be on it.\n> Do one of the following:\n>>> Go on the Fifth Quest if The Questing Beast is NOT present.\n>>> Defeat The Questing Beast by declaring as Pelinor on the Fifth Quest while the Questing Beast IS present.\n>>> You MUST declare BEFORE the cards are read.\n>>> Beware, though! If The Questing Beast is not on the Fifth Quest when you declare as Pelinor, you lose and The Questing Beast wins instead.\n\nABOUT THE QUESTING BEAST:\n> The Questing Beast can see who you are.\n> The Questing Beast must play a \"The Questing Beast Was Here\" card at least once to win, but may play a \"Reversal\" card once per game.\n> If The Questing Beast does not play a \"The Questing Beast Was Here\" card at least once before the Fifth Quest, you automatically win by attending the Fifth Quest, even if The Questing Beast is present.',
        'The Questing Beast' : 'You are Neutral in this battle and have no allies in this game.\n\nYour nemesis is Pelinor, who is also Neutral.\n\nCARDS YOU CAN PLAY:\n> \"The Questing Beast Was Here.\"\n> \"Reversal\" (Only Once Per Game)\n\n\nTO WIN:\n> The Fifth Quest Must Occur.\n> You must play at least one \"The Questing Beast Was Here\" card.\n> Complete one of the following two options:\n>>> Go on the Fifth Quest undetected.\n>>> Trick Pelinor into declaring while you are NOT on the Fifth Quest.\n\nABOUT PELINOR:\n> Pelinor cannot see you, though you can see him.\n>Pelinor also wants to reach the Fifth Quest and must go on it to win.\n> Beware! If Pelinor suspects you are on the Fifth Quest, he may declare as Pelinor, causing you to lose. (If Pelinor declares incorrectly, you automatically win and Pelinor loses.)\n> If niether you nor Pelinor are on the Fifth Quest, you both lose.',
        'Kay' : 'You are neutral and equally pulled to the Good and Evil sides, but you do have one ally who is either Good or Evil. You must determine if this ally is Good or Evil and assist as best you can. You may play Success or Failure cards on missions. You only win the game if your ally wins the game. Niether Evil players nor Merlin can identify you as Good or Evil.',
}.get(role,'ERROR: No description available.')

# get_role_information: this is called to populate information files
# blank roles:
# - Lancelot: no information
# - Arthur: no information
# - Guinevere: too complicated to generate here
# - Colgrevance: name, role (evil has an update later to inform them about the presence of Colgrevance)
def get_role_information(my_player,players,relics):
    return {
        'Tristan' : [[f'{relic.decoy2}' for relic in relics if relic.name == 'Excalibur'], ['{} is Iseult.'.format(player.name) for player in players if player.role == 'Iseult']],
        'Iseult' : [[f'{relic.decoy1}' for relic in relics if relic.name == 'Excalibur'], ['{} is Tristan.'.format(player.name) for player in players if player.role == 'Tristan']],
        'Merlin' : ['{} is Evil'.format(player.name) for player in players if (player.team == 'Evil' and player.role != 'Mordred') or player.role == 'Dagonet'],
        'Percival' : ['{} is Merlin or Morgana.'.format(player.name) for player in players if player.role == 'Merlin' or player.role == 'Morgana'],
        'Arthur' : [f'{player.name} is seeking Excalibur in the correct location.' for relic in relics if relic.name == 'Excalibur' for role in relic.location_seeker for player in players if player.role == role],
        'Titania' : [],
        'Nimue' : [],
        'Galahad' : [],
        'Guinevere' : [str(get_rumors(my_player, players,relics))],
        'Lamorak' : [str(get_relationships(my_player, players))],
        'Ector' : [f'{player.role} is in this game.' for player in players if player.team == 'Good' and player.role != 'Ector'],
        'Dagonet' : ['{} is Arthur.'.format(player.name) for player in players if player.role == 'Arthur'],
        'Uther' : [],
        'Bedivere' : [],
        'Gawain' : [f'{player.name} is a member of the Fairy Court.' for player in players if player.role == 'Nimue' or player.role == 'Bertilak' or player.role == 'Mab' or player.role == 'Oberon'],
        'Bors' : [],
        'Bertilak' : [],

        'Mordred' : ['{} is Evil.'.format(player.name) for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')],
        'Morgana' : ['{} is Evil.'.format(player.name) for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')],
        'Maleagant' : ['{} is Evil.'.format(player.name) for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')],
        'Agravaine' : ['{} is Evil.'.format(player.name) for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')],
        'Colgrevance' : ['{} is {}.'.format(player.name, player.role) for player in players if player.team == 'Evil' and player != my_player],
        'Accolon' : [[f'{player.name} is Arthur.' for player in players if player.role == 'Arthur'], [f'{player.name} is Evil.' for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')]],
        'Lancelot' : [f'{player.name} is Evil.' for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')],
        'Vortigurn' : [f'{player.name} is Evil.' for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')],
        'Annowre' : [[f'{player.name} is Evil.' for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')], [f'{relic.location}' for relic in relics if relic.name == 'Excalibur']],
        'Palamedes' : [[f'{player.name} is Evil.' for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')], [f'{player.name} is Iseult.' for player in players if player.role == 'Iseult']],
        'Mab' : [f'{player.name} is Evil.' for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')],
        'Oberon' : [[f'{player.name} is Evil.' for player in players if (player.team == 'Evil' and player != my_player and player.role != 'Colgrevance') or (player.role == 'Titania') or (player.role == 'Dagonet')], [f'{player.name} is Gawain.' for player in players if player.role == 'Gawain'], [f'{player.name} is Bertilak.' for player in players if player.role == 'Bertilak']],

        'Pelinor' : [],
        'The Questing Beast' : ['{} is Pelinor.'.format(player.name) for player in players if player.role == 'Pelinor'],
        'Kay' : [str(get_ally(my_player, players))],
    }.get(my_player.role,[])

def get_role_victory_points(role):
    return {
        # Good Roles
        'Arthur' : 'If you expended Excalibur\'s power, all other Good players gain 1 Victory Point and you gain 2 Victory Points.\n       If you do not obtain Excalibur, you lose 1 Victory Point.',
        'Bedivere' : 'If you used your Suspend power, gain 1 Victory Point.\n',
        # 'Bertilak' : 'NONE',
        'Bors' : 'For each time you used Cleanse to negate a Rot or Regrowth, gain 1 Victory Point.\n      For each time you used Cleanse and did not negate a Rot or Regrowth, lose 1 Victory Point.',
        # 'Dagonet' : 'NONE',
        # 'Ector' : 'NONE',
        'Galahad' : 'If you used your Reveal power, gain 1 Victory Point.\n',
        'Gawain' : 'If you are assassinated, you lose 2 Victory Points, all Evil players gain 1 Victory Point, and all Fae players gain 4 Victory Points.\n',
        'Guinevere' : 'If you are assassinated, you lose 2 Victory Points and all Evil players gain 2 Victory Points.\n',
        'Iseult' : 'If you are assassinated as Iseult, you lose 4 Victory Points and all Evil players gain 4 Victory Points.\n        If you are assassinated as a Lover (and not named as Iseult), you lose 3 Victory Points and all Evil players gain 3 Victory Points.\n        If Tristan is Assassinated, you lose 2 Victory Points.\n',
        'Lamorak' : 'If you are assassinated, you lose 3 Victory Points and all Evil players gain 3 Victory Points.\n',
        'Merlin' : 'If you are assassinated, you lose 4 Victory Points and all Evil players gain 4 Victory Points.\n',
        'Nimue' : 'If Excalibur is in the Lake at the end of the game, you gain 1 Victory Point.\n',
        'Percival' : 'If Merlin is assassinated, you lose 3 Victory Points.\n',
        # 'Titania' : 'NONE',
        'Tristan' : 'If you are assassinated as Tristan, you lose 4 Victory Points and all Evil players gain 4 Victory Points.\n        If you are assassinated as a Lover (and not named as Tristan), you lose 3 Victory Points and all Evil players gain 3 Victory Points.\n        If Iseult is Assassinated, you lose 2 Victory Points.\n'',
        'Uther' : 'If you used your Exile power, gain 1 Victory Point.\n',

        # Evil Roles
        'Accolon' : 'If you expended Excalibur\'s power, all other Evil players gain 1 Victory Point and you gain 2 Victory Points.\n       If you do not obtain Excalibur, you lose 1 Victory Point.',
        # 'Agravaine' : 'NONE',
        'Annowre' : 'If are in possession of the Excalibur card for the correct location at the end of the game, you gain 3 Victory Points.\n        Each time you are caught taking an Excalibur card, you must return it to the Excalibur stack and you lose 1 Victory Point.\n        If Excalibur\'s power is expended, you lose 5 Victory Points.\n',
        # 'Colgrevance' : 'NONE',
        'Lancelot' : 'If the assassination attempt is successful at all, you lose 3 Victory Points.\n          If you are assassinated while a valid target, you lose 3 additional Victory Points.\n',
        # 'Mab' : 'NONE',
        'Maleagant' : 'Each time the Scabbard is given to you, you gain 1 Victory Point.\n           If the Scabbard is never given to you, you lose 2 Victory Points.\n',
        'Mordred' : 'At the end of the game, before the Assassination Attempt, choose one player. If that player loses one or more Victory Points, you may gain the same number of Victory Points.\n         If your chosen player does not lose any Victory Points, that player gains 2 Victory Points and you lose 2 Victory Points.\n',
        # 'Morgana' : 'NONE',
        # 'Oberon' : 'NONE',
        'Palamedes' : 'If you Exiled Tristan, you gain 3 Victory Points.\n           If you Exiled a player that is not Tristan, you lose 1 Victory Point.\n           If Tristan or Iseult are assassinated, you lose 6 Victory Points.\n',
        #'Vortigurn' : 'NONE',

        # Neutral Roles
        'Kay' : 'Your Victory Point total is equal to the Victory Point total of your ally.',
        'Pelinor' : 'If you attended the Fifth Quest, you gain 2 Victory Points.\n         If you declared as Pelinor while the Questing Beast was on the Fifth Quest with you, you gain 3 Victory Points.\n         If you declared as Pelinor while the Questing Beast was not on the Fifth Quest with you, you lose 5 Victory Points.\n         If no The Questing Beast Was Here cards were played before the Fifth Quest, you gain 5 Victory Points.\n',
        'The Questing Beast' : 'If you DID attend the Fifth Quest and Pelinor did NOT declare, you gain 5 Victory Points.\n                    If you DID attend the Fifth Quest and Pelinor DID declare, you lose 3 Victory Points.\n                    If you did NOT attend the Fifth Quest and Pelinor DID declare, you gain 5 Victory Points.\n                    If you did NOT attend the Fifth Quest and Pelinor did NOT declare, you lose 3 Victory Points.\n                    If no The Questing Beast Was Here cards were played before the Fifth Quest, you lose 5 Victory Points.\n',
}.get(role,'ERROR: No description available.')

get_playable_cards(role):
    return {
        # Good Roles
        'Arthur' : 'Success\n',
        'Bedivere' : 'Success\n',
        'Bertilak' : 'Regrowth\n',
        'Bors' : 'Success\nCleanse\n',
        'Dagonet' : 'Success\n',
        'Ector' : 'Success\n',
        'Galahad' : 'Success\n',
        'Gawain' : 'Success\n',
        'Guinevere' : 'Success\n',
        'Iseult' : 'Success\n',
        'Lamorak' : 'Success\n',
        'Merlin' : 'Success\n',
        'Nimue' : 'Regrowth\n',
        'Percival' : 'Success\n',
        'Titania' : 'Success\n',
        'Tristan' : 'Success\n'',
        'Uther' : 'Success\n',

        # Evil Roles
        'Accolon' : 'Success\nFailure\n',
        'Agravaine' : 'Success\nFailure\n',
        'Annowre' : 'Success\nFailure\n',
        'Colgrevance' : 'Success\nFailure\n',
        'Lancelot' : 'Success\nFailure\n',
        'Mab' : 'Rot\n',
        'Maleagant' : 'Failure\n',
        'Mordred' : 'Success\nFailure\n',
        'Morgana' : 'Success\nFailure\n',
        'Oberon' : 'Rot\nRegrowth\n',
        'Palamedes' : 'Success\nFailure\n',
        'Vortigurn' : 'Success\nFailure\n',

        # Neutral Roles
        'Kay' : 'Success\nFailure\n',
        'Pelinor' : 'Success\nReversal\n',
        'The Questing Beast' : 'The Questing Beast Was Here\n',
}.get(role,'ERROR: No description available.')

get_conditional_cards(role):
    return {
        # Good Roles
        # 'Arthur' : 'NONE',
        'Bedivere' : '[The Suspended Card] (Must be played the next time you attend a quest.)\n',
        'Bertilak' : 'Rot (Must be played when the Holy Grail is played on you.)\n',
        'Bors' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        # 'Dagonet' : 'NONE',
        # 'Ector' : 'NONE',
        # 'Galahad' : 'NONE',
        'Gawain' : 'Reversal (If there are no Fae players in this game.)',
        # 'Guinevere' : 'NONE',
        # 'Iseult' : 'NONE',
        # 'Lamorak' : 'NONE',
        'Merlin' : 'Reversal (Available when Excalibur is in the Lake, unless the Holy Grail is currently played on you.)\n',
        'Nimue' : 'Success (Must be played when the Holy Grail is played on you, unless Excalibur is in the Lake.',
        'Percival' : 'Reversal (Available when Excalibur is at Camelot, unless the Holy Grail is currently played on you.)',
        # 'Titania' : 'NONE',
        # 'Tristan' : 'NONE'',
        # 'Uther' : 'NONE',

        # Evil Roles
        'Accolon' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        # 'Agravaine' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        'Annowre' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        # 'Colgrevance' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        'Lancelot' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        # 'Mab' : 'NONE',
        'Maleagant' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        'Mordred' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        'Morgana' : 'Success (Must be played when the Holy Grail is played on you.)\nReversal (Available when Excalibur is in the Stone, unless the Holy Grail is currently played on you.)\n',
        'Oberon' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        'Palamedes' : 'Success (Must be played when the Holy Grail is played on you.)\nSuccess (Must be played on each quest after you have Exiled Tristan.)',
        'Vortigurn' : 'Success (Must be played when the Holy Grail is played on you.)\n',

        # Neutral Roles
        'Kay' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        'Pelinor' : 'Success (Must be played when the Holy Grail is played on you.)\n',
        'The Questing Beast' : 'Reversal (Available once per game.)\nSuccess (Must be played when the Holy Grail is played on you.)\n',
}.get(role,'ERROR: No description available.')

def get_rumors(my_player, players, relics):
    rumors = []

    if my_player.role != 'Guinevere':
        return
    
    # Generate rumors about Merlin
    merlin_player = None
    is_Merlin = 0
    for player in players:
        if player.role == 'Merlin':
            merlin_player = player.name
            is_Merlin = 1
    if is_Merlin == 1:
        for player in players:
            if (player.team == 'Evil' and player.role != 'Mordred') or player.role == "Dagonet":
                player_of_evil = player.name
                rumors.append('{} sees {}'.format(merlin_player, player_of_evil))

    # Generate rumors about Percival
    percival_player = None
    is_Percival = 0
    for player in players:
        if player.role == 'Percival':
            percival_player = player.name
            is_Percival = 1
    if is_Percival == 1:
        for player in players:
            if player.role == 'Merlin' or player.role == 'Morgana':
                seer = player.name
                rumors.append('{} sees {}'.format(percival_player, seer))

    # Generate rumor about the Lovers
    tristan_player = None
    iseult_player = None
    is_Lovers = 0
    for player in players:
        if player.role == 'Tristan':
            tristan_player = player.name
            is_Lovers += 1
        elif player.role == 'Iseult':
            iseult_player = player.name
            is_Lovers += 1
    if is_Lovers == 2:
        rumors.append('{} sees {}'.format(tristan_player, iseult_player))
        rumors.append('{} sees {}'.format(iseult_player, tristan_player))

    # Generate rumor about Evil players
    for player in players:
        if player.team == 'Evil' and player.role != 'Mordred':
            for player_two in players:
                if (player_two.team == 'Evil' and player_two.role != 'Mordred' and player_two.role != 'Colgrevance' and player_two != player) or (player_two.role == 'Titania' and player_two != player):
                    rumors.append('{} sees {}'.format(player.name, player_two.name))

    # Generate rumor about Ector
    is_Ector = 0
    for player in players:
        if player.role == 'Ector':
            is_Ector = 1
    if is_Ector == 1:
        for player in players:
            if player.team == 'Good' and player.role != 'Ector' and player.role != 'Guinevere':
                rumors.append(f'Sir Ector sees {player.role}')

    # Generate rumor about The Questing Beast
    questing_player = None
    is_Questing = 0
    for player in players:
        if player.role == 'The Questing Beast':
            questing_player = player
            is_Questing = 1
    if is_Questing == 1:
        for player in players:
            if player.role == 'Pelinor':
                rumors.append(f'{questing_player} sees {player}.')

    # Generate rumors about Arthur
    arthur_player = None
    is_Arthur = 0
    for player in players:
        if player.role == 'Arthur':
            is_Arthur = 1
            arthur_player = player
    for relic in relics:
        if relic == 'Excalibur':
            for role in relic.location_seeker:
                for player in players:
                    if player.role == role:
                        rumors.append(f'{arthur_player.name} sees {player.name}')

    # Pick two rumors and return them as a string.
    rumor_one = random.choice(rumors)
    rumor_two = random.choice(rumors)
    while rumor_one == rumor_two:
            rumor_two = random.choice(rumors)
    return rumor_one + '\n' + rumor_two

def get_relationships(my_player, players):

    if my_player.role != 'Lamorak':
        return
    
    # Assign teams
    good_team = []
    evil_team = []
    neutral_team = []
    for player in players:
        if player.team == 'Good' and player.role != 'Lamorak':
            good_team.append(player)
        if player.team == 'Evil' and player.role != 'Mordred':
            evil_team.append(player)
        if player.team == 'Neutral' and player.role != 'Kay':
            neutral_team.append(player)
    valid_players = good_team + evil_team + neutral_team

    # Choose random Opposing Team players
    opposition = None
    opposing_player = random.choice(valid_players)
    if opposing_player.team == 'Good':
            opposition = opposing_player.name + ' opposes ' + (random.choice(evil_team)).name
    elif opposing_player.team == 'Evil':
            opposition = opposing_player.name + ' opposes ' + (random.choice(good_team)).name
    elif opposing_player.role == 'Pelinor':
            for player in players:
                if player.role == 'The Questing Beast':
                    opposition = opposing_player.name + ' opposes ' + player.name
    elif opposing_player.role == 'The Questing Beast':
            for player in players:
                if player.role == 'Pelinor':
                    opposition = opposing_player.name + ' opposes ' + player.name

    # Choose random Collaborator players
    # Random choice of good or evil team (else good would be much more likely)
    # while loop to prevent collaborators from being the same player
    collaboration = None
    player_one = "1"
    player_two = "1"
    random_team = random.choice(['Good','Evil'])
    while player_one == player_two:
        if random_team == 'Good':
            player_one = (random.choice(good_team)).name
            player_two = (random.choice(good_team)).name
        elif random_team == 'Evil':
            player_one = (random.choice(evil_team)).name
            player_two = (random.choice(evil_team)).name

    collaboration = player_one + ' is collaborating with ' + player_two
    return opposition + '\n' + collaboration


def get_ally(my_player, players):
    
    if my_player.role != 'Kay':
        return
                    
    # Get Kay's team and make a list of players on that team.
    allies = []
    if my_player.secret == 'Good':
        for player in players:
            if player.team == 'Good':
                allies.append(player)
    if my_player.secret == 'Evil':
        for player in players:
            if player.team == 'Evil':
                allies.append(player)

    # Return a random ally.
    kay_ally = random.choice(allies)
    return f'{kay_ally.name} is your ally. If {kay_ally.name} wins the game, so do you.'


# EXCALIBUR
# Randomly choose a location for Excalibur and keep track of decoy locations.
def get_excalibur():
    excalibur_hiding_places = [' in the Stone', ' at Camelot', ' in the Lake']

    # Randomly select one location for Excalibur to be
    excalibur_location = random.choice(excalibur_hiding_places)
    excalibur_hiding_places.remove(excalibur_location)

    # Produce the two decoy locations.
    excalibur_decoy1 = random.choice(excalibur_hiding_places)
    excalibur_hiding_places.remove(excalibur_decoy1)
    excalibur_decoy2 = random.choice(excalibur_hiding_places)

    return excalibur_location, excalibur_decoy1, excalibur_decoy2
          

class Player():
    # Players have the following traits
    # name: the name of the player as fed into system arguments
    # role: the role the player possesses
    # team: whether the player is on good, evil, or neutral's team
    # type: information or ability
    # seen: a list of what they will see
    # modifier: the random modifier this player has [NOT CURRENTLY UTILIZED]
    def __init__(self, name):
        self.name = name
        self.role = None
        self.team = None
        self.origin = None
        self.modifier = None
        self.info = []
        self.is_assassin = False
        self.secret = None

    def set_role(self, role):
        self.role = role

    def set_team(self, team):
        self.team = team
        
    def set_origin(self, origin):
        self.origin = origin

    def add_info(self, info):
        self.info += info

    def erase_info(self, info):
        self.info = []

    def generate_info(self, players):
        pass

class Relic():
    # Relics have the following traits
    # name: the name of the relic as fed into system arguments
    # type: the type the relic possesses
    # location: where the relic may be claimed
    # decoy1: first location where the relic is not found
    # decoy2: second location where the relic is not found
    # location_seeker: list of roles that seek the same location as the relic
    # decoy1_seeker: list of roles that seek the same location as decoy1
    # decoy2_seeker:list of roles that seek the same location as decoy2
    def __init__(self, name):
        self.name = name
        self.role = None
        self.location = None
        self.decoy1 = None
        self.decoy2 = None
        self.location_seeker = []
        self.decoy1_seeker = []
        self.decoy2_seeker = []

    def set_type(self, type):
        self.type = type

    def set_seeker(self, location):
        if location == ' in the Stone':
            return 'Uther', 'Vortigurn', 'Morgana', 'Palamedes', 'Sebile'
        elif location == ' at Camelot':
            return 'Percival', 'Galahad', 'Lancelot'
        elif location == ' in the Lake':
            return 'Merlin', 'Bedivere', 'Agravaine', 'Nimue'

    def set_location(self, location):
        seekers = self.set_seeker(location)
        self.location = 'Excalibur is' + location
        for seeker in seekers:
            self.location_seeker.append(seeker)

    def set_decoy1(self, decoy1):
        seekers = self.set_seeker(decoy1)
        self.decoy1 = 'Excalibur is not' + decoy1
        for seeker in seekers:
            self.decoy1_seeker.append(seeker)

    def set_decoy2(self, decoy2):
        seekers = self.set_seeker(decoy2)
        self.decoy2 = 'Excalibur is not' + decoy2
        for seeker in seekers:
            self.decoy2_seeker.append(seeker)

def get_player_info(player_names):
    num_players = len(player_names)
    if len(player_names) != num_players:
        print('ERROR: Duplicate player names.')
        exit(1)

    # Place Excalibur and decoy locations.
    relics = []
    excalibur_info = get_excalibur()
    excalibur = Relic("Excalibur")
    relics.append(excalibur)
    excalibur.set_type('Sword')
    excalibur.set_location(excalibur_info[0])
    excalibur.set_decoy1(excalibur_info[1])
    excalibur.set_decoy2(excalibur_info[2])

    # create player objects
    players = []
    for i in range(0, len(player_names)):
        player = Player(player_names[i])
        players.append(player)

    # number of good and evil roles
    num_neutral = 0
    if num_players <= 6:
        if random.choice([True, False, False]):
            kay_team = random.choice(['Good','Evil'])
            if kay_team == 'Good':
                num_evil = 2
                num_neutral = 1
            elif kay_team == 'Evil':
                num_evil = 1
                num_neutral = 1
        else:
            num_evil = 2
    elif num_players <= 8:
        if random.choice([True, False, False, False]):
            kay_team = random.choice(['Good','Evil'])
            if kay_team == 'Good':
                num_evil = 3
                num_neutral = 1
            elif kay_team == 'Evil':
                num_evil = 2
                num_neutral = 1
        else:
            num_evil = 3
    elif num_players <= 10:
        choice = random.choice(['kay', 'hunt', 'kayandhunt', 'niether', 'niether'])
        if choice == 'hunt':
            num_evil = 3
            num_neutral = 2
        elif choice == 'kay':
            kay_team = random.choice(['Good','Evil'])
            if kay_team == 'Good':
                num_evil = 4
                num_neutral = 1
            elif kay_team == 'Evil':
                num_evil = 3
                num_neutral = 1
        elif choice == 'kayandhunt':
            kay_team = random.choice(['Good','Evil'])
            if kay_team == 'Good':
                num_evil = 3
                num_neutral = 3
            elif kay_team == 'Evil':
                num_evil = 2
                num_neutral = 3
        elif choice == 'niether':
            num_evil = 4
    elif num_players <= 12:
        choice = random.choice(['kay', 'hunt', 'kayandhunt', 'niether', 'niether'])
        if choice == 'hunt':
            num_evil = 4
            num_neutral = 2
        elif choice == 'kay':
            kay_team = random.choice(['Good','Evil'])
            if kay_team == 'Good':
                num_evil = 5
                num_neutral = 1
            elif kay_team == 'Evil':
                num_evil = 4
                num_neutral = 1
        elif choice == 'kayandhunt':
            kay_team = random.choice(['Good','Evil'])
            if kay_team == 'Good':
                num_evil = 4
                num_neutral = 3
            elif kay_team == 'Evil':
                num_evil = 3
                num_neutral = 3
        elif choice == 'niether':
            num_evil = 5
    num_good = num_players - num_evil - num_neutral

    # establish available roles
    good_roles = ['Merlin', 'Percival', 'Arthur', 'Dagonet']
    evil_roles = ['Mordred', 'Morgana', 'Accolon', 'Annowre']
    neutral_roles = ['Kay']
                    
    # 6 plus
    if num_players > 5:
        good_roles.extend(['Tristan', 'Iseult', 'Guinevere'])
        evil_roles.extend(['Palamedes', 'Maleagant'])

    # 7 plus
    if num_players > 6:
        good_roles.extend(['Titania', 'Uther', 'Bedivere', 'Galahad'])
        evil_roles.extend(['Vortigurn', 'Lancelot', 'Agravaine'])

    # 8 plus
    if num_players > 7:
        good_roles.extend(['Nimue', 'Bertilak', 'Gawain', 'Bors', 'Ector'])
        evil_roles.extend(['Mab', 'Oberon'])

    # 10 plus
    if num_players > 9:
        good_roles.append('Lamorak')
        evil_roles.append('Colgrevance')
        neutral_roles.extend(['Pelinor', 'The Questing Beast'])

    good_roles_in_game = random.sample(good_roles, num_good)
    evil_roles_in_game = random.sample(evil_roles, num_evil)
    neutral_roles_in_game = []
    if num_neutral == 1:
        neutral_roles_in_game.append('Kay')
    elif num_neutral == 2:
        neutral_roles_in_game.extend(['Pelinor', 'The Questing Beast'])
    elif num_neutral == 3:
        neutral_roles_in_game.extend(['Kay', 'Pelinor', 'The Questing Beast'])
        

    # lone lovers are rerolled
    # 50% chance to reroll one lone lover
    # 50% chance to reroll another role into a lover
    if sum(gr in ['Tristan','Iseult'] for gr in good_roles_in_game) == 1 and num_good > 1:
        if 'Tristan' in good_roles_in_game:
            good_roles_in_game.remove('Tristan')
        if 'Iseult' in good_roles_in_game:
            good_roles_in_game.remove('Iseult')

        if random.choice([True, False]):
            # replacing the lone lover
            available_roles = good_roles
            not_available = good_roles_in_game + ['Tristan'] + ['Iseult']
            for role in not_available:
                available_roles.remove(role)
            # DecrecationWarning issue. Found solution at https://stackoverflow.com/questions/70426576/get-random-number-from-set-deprecation
            good_roles_in_game.append(random.choice(available_roles))
        else:
            # upgradng to pair of lovers
            rerolled = random.choice(good_roles_in_game)
            good_roles_in_game.remove(rerolled)
            good_roles_in_game.append('Tristan')
            good_roles_in_game.append('Iseult')

    # Add Palamedes to the game. 50% chance if Tristan and Iseult are in the game.
    is_Palamedes = False
    if (sum(gr in ['Tristan','Iseult'] for gr in good_roles_in_game) == 2) and (num_good > 2):
        for er in evil_roles_in_game:
            if er == 'Palamedes':
                is_Palamedes = True
                break
            else:
                continue
        if is_Palamedes == False and random.choice([True,False]):
            rerolled = random.choice(evil_roles_in_game)
            evil_roles_in_game.remove(rerolled)
            evil_roles_in_game.append('Palamedes')

    # Ensure there is a role that can be the Assassin.
    if num_evil = 1 and (sum(er in ['Lancelot','Palamedes', 'Oberon'] for er in evil_roles_in_game) >= 1):
        incompatable_removed = 0
        for roles in evil_roles_in_game:
            if incompatable_removed == 0 and (role == 'Lancelot' or role == 'Palamedes' or role == 'Oberon'):
                evil_roles_in_game.remove(role)
                incompatable_removed += 1
        available_roles = evil_roles
        not_available = evil_roles_in_game + ['Lancelot'] + ['Palemedes'] + ['Oberon']
        for role in not_available:
            available_roles.remove(role)
        evil_roles_in_game.append(random.choice(available_roles))
    elif num_evil = 2 and (sum(er in ['Lancelot','Palamedes', 'Oberon'] for er in evil_roles_in_game) >= 2):
        incompatable_removed = 0
        for roles in evil_roles_in_game:
            if incompatable_removed == 0 and (role == 'Lancelot' or role == 'Palamedes' or role == 'Oberon'):
                evil_roles_in_game.remove(role)
                incompatable_removed += 1
        available_roles = evil_roles
        not_available = evil_roles_in_game + ['Lancelot'] + ['Palemedes'] + ['Oberon']
        for role in not_available:
            available_roles.remove(role)
        evil_roles_in_game.append(random.choice(available_roles))
    elif num_evil = 3 and (sum(er in ['Lancelot','Palamedes', 'Oberon'] for er in evil_roles_in_game) >= 3):
        incompatable_removed = 0
        for roles in evil_roles_in_game:
            if incompatable_removed == 0 and (role == 'Lancelot' or role == 'Palamedes' or role == 'Oberon'):
                evil_roles_in_game.remove(role)
                incompatable_removed += 1
        available_roles = evil_roles
        not_available = evil_roles_in_game + ['Lancelot'] + ['Palemedes'] + ['Oberon']
        for role in not_available:
            available_roles.remove(role)
        evil_roles_in_game.append(random.choice(available_roles))
            
            
    # roles after validation
    #print(good_roles_in_game)
    #print(evil_roles_in_game)

    # role assignment
    random.shuffle(players)
    neutral_players = []
    good_players = players[:num_good]
    if num_neutral == 0:
        evil_players = players[num_good:]
    else:
        goodAndEvil = num_good + num_evil
        evil_players = players[num_good:goodAndEvil]
        neutral_players = (players[goodAndEvil:])

    player_of_role = dict()

    # Good Team: Set Team and Origin.
    for gp in good_players:
        new_role = good_roles_in_game.pop()
        gp.set_role(new_role)
        gp.set_team('Good')
        player_of_role[new_role] = gp
        if gp.role == 'Bertilak' or gp.role == 'Nimue':
            gp.set_origin('Fae')
        else:
            gp.set_origin('Mortal')

    # Assign Compatable Evil Player to Assassin role.
    if evil_players[0].role != 'Lancelot' and evil_players[0].role != 'Palamedes' and evil_players[0].role != 'Oberon':
        evil_players[0].is_assassin = True
    elif evil_players[1].role != 'Lancelot' and evil_players[0].role != 'Palamedes' and evil_players[0].role != 'Oberon':
        evil_players[1].is_assassin = True
    elif evil_players[2].role != 'Lancelot' and evil_players[0].role != 'Palamedes' and evil_players[0].role != 'Oberon':
        evil_players[2].is_assassin = True
    else:
        evil_players[3].is_assassin = True

    # Evil Team: Set Team and Origin.
    for ep in evil_players:
        new_role = evil_roles_in_game.pop()
        ep.set_role(new_role)
        ep.set_team('Evil')
        player_of_role[new_role] = ep
        if ep.role == 'Mab' or ep.role == 'Oberon':
            ep.set_origin('Fae')
        else:
            ep.set_origin('Mortal')

    # Neutral Team: Set Team and Origin.
    for np in neutral_players:
        new_role = neutral_roles_in_game.pop()
        np.set_role(new_role)
        np.set_team('Neutral')
        player_of_role[new_role] = np
        np.set_origin('Mortal')
                    
    for player in players:
        if player.role == 'Kay':
            player.secret = kay_team
            kay_name = player.name
            kay_player = player
    
    for p in players:
        p.add_info(get_role_information(p,players,relics))
        try:
            if isinstance(p.info[0], list):
                try:
                    p.info = p.info[0] + p.info[1]
                except Exception:
                    pass
        except Exception:
            pass
        random.shuffle(p.info)
        # print(p.name,p.role,p.team,p.info)

    # Informing Evil about Colgrevance
    for ep in evil_players:
        if ep.role != 'Colgrevance' and player_of_role.get('Colgrevance'):
            ep.add_info(['Colgrevance lurks in the shadows. (There is another Evil that you do not see.)'])
        if ep.role != 'Colgrevance' and player_of_role.get('Titania'):
            ep.add_info(['Titania has infiltrated your ranks. (One of the people you see is not Evil.)'])
        if ep.is_assassin:
            ep.add_info(['You are the Assassin.'])

    # delete and recreate game directory
    if os.path.isdir("game"):
        shutil.rmtree("game")
    os.mkdir("game")

    # for player in players:
    #     run = document.add_paragraph().add_run('some text')
    #     font = run.font
    #     font.color.rgb = RGBColor(0x42, 0x24, 0xE9)
    #     document.add_paragraph(player.name + ', you are' + player.role + '.')

    bar= '----------------------------------------\n'
    for player in players:
        player.string= bar+'You are '+player.role+' ['+player.team+']\n'+bar+get_role_description(player.role)+'\n'+bar+'\n'.join(player.info)+'\n'+bar+''.join(get_victory_points(player, players))+bar
        player_file = "game/{}.docx".format(player.name)
        with open(player_file,"w") as file:
            file.write(player.string)
########################################################## DOCUMENT REWRITE

    document = Document()
    sections = document.sections
    section = sections[0]

    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    paragraph = document.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = 0
    paragraph_format.space_after = 0

    ####### FIX INTENTATION BELOW:

run = paragraph.add_run("     "+player.name+",")
font = run.font
font.name = 'Breathe Fire III'
font.size = Pt(50)

paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0

run = paragraph.add_run("                            you are ")
font = run.font
font.name = 'Breathe Fire III'
font.size = Pt(30)


if player.team == 'Good':
    run = paragraph.add_run(player.role+"\n")
    font = run.font
    font.name = 'Breathe Fire III'
    font.size = Pt(30)
    font.color.rgb = RGBColor(0, 255, 0) # Green
elif player.team == 'Evil':
    run = paragraph.add_run(player.role+"\n")
    font = run.font
    font.name = 'Breathe Fire III'
    font.size = Pt(30)
    font.color.rgb = RGBColor(255, 0, 0) # Red
elif player.team == 'Neutral':
    run = paragraph.add_run(player.role+"\n")
    font = run.font
    font.name = 'Breathe Fire III'
    font.size = Pt(30)
    font.color.rgb = RGBColor(255, 215, 0) # Gold

insertHR(paragraph)


# Allegiance
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0

run = paragraph.add_run("\n     Allegiance: ")
font = run.font
font.name = 'Caladea'
font.size = Pt(14)
font.bold = True

if player.team == 'Good':
    run = paragraph.add_run("Good")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
if player.team == 'Evil':
    run = paragraph.add_run("Evil")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
if player.team == 'Neutral':
    run = paragraph.add_run("Neutral")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)


# Origins
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0

run = paragraph.add_run("     Origins: ")
font = run.font
font.name = 'Caladea'
font.size = Pt(14)
font.bold = True

if player.origin == "Mortal:
    run = paragraph.add_run("Mortal")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
if player.origin == "Fae:
    run = paragraph.add_run("Fae")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)

# Assassination
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0

if player.role == 'Lancelot':
    run = paragraph.add_run("     You are a CONDITIONAL ASSASSINATION TARGET. \(You may be assassinated if fewer than two quests fail.\)\n ")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.italic = True
elif player.team == 'Evil' and player.is_assassin = False:
    run = paragraph.add_run("     You are NOT the Assassin.\n ")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.italic = True
elif player.is_assassin = True:
    run = paragraph.add_run("     You are the ASSASSIN.\n ")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.italic = True
elif player.role == 'Tristan' or player.role == 'Iseult' or player.role == Merlin or player.role == 'Guinevere' or player.role == 'Lamorak' or player.role == 'Gawain':
    run = paragraph.add_run("     You are a TARGET for the Assassin.\n ")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.italic = True
else:
    run = paragraph.add_run("     You are NOT a target for the Assassin.\n ")
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.italic = True

insertHR(paragraph)

# Abilities
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0


run = paragraph.add_run("Abilities:\n ")
font = run.font
font.name = 'Caladea'
font.size = Pt(14)
font.bold = True

run = paragraph.add_run(get_role_description(player.role) + '\n')
font = run.font
font.name = 'Caladea'
font.size = Pt(14)

# Cards You Can Play
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0


run = paragraph.add_run("Cards You Can Play On Quests:\n ")
font = run.font
font.name = 'Caladea'
font.size = Pt(14)
font.bold = True

run = paragraph.add_run(get_playable_cards(player.role) + '\n')
font = run.font
font.name = 'Caladea'
font.size = Pt(14)

# Conditional Cards
try:
    paragraph = document.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = 0
    paragraph_format.space_after = 0
    run = paragraph.add_run('Conditional Cards You Can Play: (When you can play them.)\n')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.bold = True
    run = paragraph.add_run(get_conditional_cards(player.role) + '\n')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
except:
    pass

insertHR(paragraph)

# Secret Information
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0


run = paragraph.add_run("Secret Information:\n ")
font = run.font
font.name = 'Caladea'
font.size = Pt(14)
font.bold = True

run = paragraph.add_run("'\n'.join(player.info)\n ")
font = run.font
font.name = 'Caladea'
font.size = Pt(14)

insertHR(paragraph)

# Victory Points
paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = 0
paragraph_format.space_after = 0


run = paragraph.add_run("Victory Points:\n ")
font = run.font
font.name = 'Caladea'
font.size = Pt(14)
font.bold = True

# Team Points
if player.team == 'Good':
    run = paragraph.add_run('Good: ')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.bold = True
    run = paragraph.add_run('If three quests succeeded, you gain 3 Victory Points.\n')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
elif player.team == 'Evil':
    run = paragraph.add_run('Evil: ')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.bold = True
    run = paragraph.add_run('If three quests failed, you gain 3 Victory Points.\n')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)

# Origin Points
if player.origin == 'Mortal and player.role != 'Kay':
    run = paragraph.add_run('Mortal: ')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.bold = True
    run = paragraph.add_run('If the Holy Grail remains Uncorrupted, you gain 1 Victory Point.\n        If you identified a Fae during The Wild Hunt, you gain 2 Victory Points.\n')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
elif player.origin == 'Fae':
    run = paragraph.add_run('Fae: ')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.bold = True
    run = paragraph.add_run('If the Holy Grail has become Corrupted, you gain 3 Victory Points.\n        For each Mortal that identified you during the Wild Hunt, you lose 1 Victory Point.\n')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)

# Assassin Point
if player.is_Assassin == True:
    run = paragraph.add_run('Assassin: ')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.bold = True
    run = paragraph.add_run('If you correctly assassinated a Target, you gain 1 Victory Point.\n')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)

# Role Points

try:
    run = paragraph.add_run(player.role + ': ')
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
    font.bold = True
    run = paragraph.add_run(get_role_victory_points(player.role))
    font = run.font
    font.name = 'Caladea'
    font.size = Pt(14)
except:
    pass
    

document.save('word.docx')

################################################################################### DOCUMENT REWRITE
    first_player = random.sample(players,1)[0]
    fae_count = 0
    with open("game/1. Read to Start.docx", "w") as file:
        file.write("FIRST LEADER:\nThe player proposing the first mission is {}.\n\n".format(first_player.name))
        file.write("THE HOLY GRAIL\n")
        for player in players:
            if player.origin == "Fae":
                fae_count += 1
            if player.role == "Gawain":
                fae_count -= 1
        if fae_count < 1:
            fae_count = 1
        file.write(f'The Holy Grail is {(fae_count*2) + 1} Fae Spells away from Corruption.')
        #file.write("\n" + second_mission_starter + " is the starting player of the 2nd round.\n")

    with open("game/Do NOT Open.docx", "w") as file:
        file.write("Player -> Role\n\n GOOD TEAM:\n")
        for gp in good_players:
            file.write("{} -> {}\n".format(gp.name, gp.role))
        file.write("\nEVIL TEAM:\n")
        for ep in evil_players:
            file.write("{} -> {}\n".format(ep.name,ep.role))
        if len(neutral_players) > 0:
            file.write("\nNEUTRAL TEAM:\n")
            for np in neutral_players:
                file.write("{} -> {}\n".format(np.name,np.role))
        file.write(f'\nEXCALIBUR:\n')
        for relic in relics:
            if relic == "Excalibur":
                file.write(f'{relic.location}')

if __name__ == "__main__":
#    if not (6 <= len(sys.argv) <= 13):
 #       print("Invalid number of players")
  #      exit(1)

    #players = sys.argv[1:]
    players = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
    num_players = len(players)
    players = set(players) # use as a set to avoid duplicate players
    players = list(players) # convert to list
    random.shuffle(players) # ensure random order, though set should already do that
    if len(players) != num_players:
        print("No duplicate player names")
        exit(1)

    get_player_info(players)
    
    
    
    
    
    
    
    
#     from docx import Document
# from docx.enum.style import WD_STYLE_TYPE
# from docx.shared import Pt, RGBColor, Inches
# from docx.oxml.ns import qn
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# import docx

# document = Document()
# sections = document.sections
# for section in sections:
#     section.top_margin = Inches(0.5)
#     section.bottom_margin = Inches(1)
#     section.left_margin = Inches(1)
#     section.right_margin = Inches(1)

# section = document.sections[0]

# sectPr = section._sectPr
# cols = sectPr.xpath('./w:cols')[0]
# cols.set(qn('w:num'), '2')

# def writedocx(content, font_name = 'Times New Roman', font_size = 12, font_bold = False, font_italic = False, font_underline = False, color = RGBColor(0, 0, 0),
#               before_spacing = 5, after_spacing = 5, line_spacing = 1.5, keep_together = True, keep_with_next = False, page_break_before = False,
#               widow_control = False, align = 'left', style = ''):
#     paragraph = document.add_paragraph(str(content))
#     paragraph.style = document.styles.add_style(style, WD_STYLE_TYPE.PARAGRAPH)
#     font = paragraph.style.font
#     font.name = font_name
#     font.size = Pt(font_size)
#     font.bold = font_bold
#     font.italic = font_italic
#     font.underline = font_underline
#     font.color.rgb = color
#     paragraph_format = paragraph.paragraph_format
#     paragraph_format.space_before = Pt(before_spacing)
#     paragraph_format.space_after = Pt(after_spacing)
#     paragraph.line_spacing = line_spacing
#     paragraph_format.keep_together = keep_together
#     paragraph_format.keep_with_next = keep_with_next
#     paragraph_format.page_break_before = page_break_before
#     paragraph_format.widow_control = widow_control
#     if align.lower() == 'left':
#         paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
#     elif align.lower() == 'center':
#         paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
#     elif align.lower() == 'right':
#         paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
#     elif align.lower() == 'justify':
#         paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
#     else:
#         paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        
# document.save('word.docx')
