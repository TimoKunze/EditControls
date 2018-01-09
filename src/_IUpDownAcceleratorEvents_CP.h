//////////////////////////////////////////////////////////////////////
/// \class Proxy_IUpDownAcceleratorEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IUpDownAcceleratorEvents interface</em>
///
/// \if UNICODE
///   \sa UpDownAccelerator, EditCtlsLibU::_IUpDownAcceleratorEvents
/// \else
///   \sa UpDownAccelerator, EditCtlsLibA::_IUpDownAcceleratorEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IUpDownAcceleratorEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IUpDownAcceleratorEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IUpDownAcceleratorEvents