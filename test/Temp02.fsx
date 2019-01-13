// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

#r "netstandard"

type HasWordAppKey =
    abstract WordAppKey : string

type Phantom1 = class end


type MyConfig =  
    | MyConfig of unit
    override x.ToString() = "*MY_CONFIG*"
            
    interface HasWordAppKey with
        member x.WordAppKey = "WordApp"

let test01 () = 
    let x = MyConfig () in (x :> HasWordAppKey).WordAppKey

let quoteOfficeKey (mystery:HasWordAppKey) : string = 
    sprintf "'%s'" mystery.WordAppKey

let quoteOfficeKeyTwo (x:HasWordAppKey) : string = 
    sprintf "'%s'" x.WordAppKey



let test02 () = 
    let x = MyConfig () in quoteOfficeKey (x :> HasWordAppKey)



let test03 () = 
    let x = MyConfig () in quoteOfficeKeyTwo x


let test04 () = 
    let x = MyConfig () in (x :> HasWordAppKey).WordAppKey


type UserResources<'reskey> = 
     val private UserResources: Map<string,obj>
     val private KeyStore: 'reskey
     
     new (reskey:'reskey, resources:Map<string,obj>) = 
        { UserResources = resources
          KeyStore = reskey }

     member x.Map
        with get() = x.UserResources

     member x.Keymap
        with get() = x.KeyStore

type WordHandle = 
    { WordApp : unit }

type WordResources = 
    | WordRes of unit
    interface HasWordAppKey with
        member x.WordAppKey = "WordApp"

let MyResources : UserResources<WordResources> = 
    let resources : WordResources = WordRes ()
    let wordApp = { WordApp = () }
    let wordKey = (resources :> HasWordAppKey).WordAppKey
    new UserResources<WordResources> (resources, Map.ofList [ wordKey, wordApp :> obj ])

let extractWord<'T when 'T :> HasWordAppKey> (resources:UserResources<'T>) : WordHandle = 
    let x = resources.Keymap
    let key = x.WordAppKey
    let o:obj = Map.find key resources.Map
    o :?> WordHandle

let testImportant () = 
    extractWord MyResources

// Actually, HasWordAppKey works so well we don't need the registry...

type WordApplication = 
    | WordApp of unit

type HasWordHandle =
    abstract WordAppHandle : WordApplication

type MyConfig2 = 
    { WordHandle : WordApplication }
    interface HasWordHandle with
        member x.WordAppHandle = x.WordHandle


let runWord<'T when 'T :> HasWordHandle> (resources:'T) : string = 
    let x:WordApplication = resources.WordAppHandle
    match x with | WordApp () -> "YES!"

let test05 () = 
    let myconfig = { WordHandle = WordApp () }
    runWord myconfig




