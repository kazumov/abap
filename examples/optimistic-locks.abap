*&---------------------------------------------------------------------*
*& Report DEMO_OPT_ENQUEUE
*&
*&---------------------------------------------------------------------*
*& Demo program for optimistic locks
*&
*& Besides the old lock modes
*& - Non cumulative lock Mode = 'X'
*& - Exclusive lock Mode = 'E'
*& - Shared lock Mode = 'S'
*&
*& a new lock mode is introduced
*& - Optimistic lock Mode = 'O'
*& - Mode 'R' for pRomoting an 'O' lock into an 'E' lock
*& Mode 'R' does not appear in the lock table
*&
*& See MODE_xxx-Parameter of the ENQUEUE_xxx FMs
*&
*& An optimistic lock is similar to a shared lock 'S'
*& The differences are the following:
*&
*& - it can be promoted to an exclusive lock (mode 'E'),
*& if there is no collision
*& The promotion of n optimistic lock collides with
*& - a shared lock 'S' which didn't collide with the 'O' lock
*& but collides with the resulting 'E' lock
*& - an own lock becomes a foreign lock due to passing it to the
*& update
*&
*& - The successful promotion of an optimistic lock to an exclusive lock
*& deletes all foreign optimistic locks which collide with the
*& generated 'E' lock.
*&
*& - the optimistic lock can get lost due to the promotion of a
*& foreign 'O' lock (see above).
*&
*& Optimistic locks are compatible to the known lock modes.
*& Programs using ordinary pessimistic locks can coexist with programs
*& that switched to optimistic locking and use the same
*& Lock Objects and Enqueue function modules.
*& You don't need to switch all programs at the same time
*& Please take care of side effects!
*&
*& The existing ENQUEUE_xxx and DEQUEUE_xxx function modules can also
*& be used for otimistic locks.
*&
*&---------------------------------------------------------------------*
REPORT DEMO_OPT_ENQUEUE.
DATA: lock_mode_opt type enqmode value 'O'.
DATA: lock_mode_excl type enqmode value 'E'.
DATA: lock_mode_opt_to_excl type enqmode value 'R'.
*-----------------------------------------------------------------------
* Set optimistic lock.
* The operation is successful when there is no collision with foreign
* 'E' and 'X' locks.
*-----------------------------------------------------------------------
CALL FUNCTION 'ENQUEUE_ENQ_SFLIGHT'
EXPORTING
MODE_SFLIGHT = lock_mode_opt
EXCEPTIONS
FOREIGN_LOCK = 1
SYSTEM_FAILURE = 2
OTHERS = 3.
IF SY-SUBRC <> 0.
MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
ENDIF.
*-----------------------------------------------------------------------
* Promote optimistic lock into exclusive lock.
* The operation is successful, if
* - the own 'O' lock still exists
* - there is no collision of the new 'E' lock with foreign locks
*-----------------------------------------------------------------------
CALL FUNCTION 'ENQUEUE_ENQ_SFLIGHT'
EXPORTING
MODE_SFLIGHT = lock_mode_opt_to_excl
EXCEPTIONS
FOREIGN_LOCK = 1
SYSTEM_FAILURE = 2
OTHERS = 3.
IF SY-SUBRC <> 0.
MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
ENDIF.
*-----------------------------------------------------------------------
* Release exclusive lock
*-----------------------------------------------------------------------
CALL FUNCTION 'DEQUEUE_ENQ_SFLIGHT'
EXPORTING
MODE_SFLIGHT = lock_mode_excl.
