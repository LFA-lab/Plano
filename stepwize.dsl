 françaisworkspace "Stepwize" {
  
  model {
    user = person "User"

    rails = softwareSystem "Rails Frontend" {
      railsApp = container "Rails App"
    }

    api = softwareSystem "FastAPI Backend" {
      fastapiService = container "FastAPI Service"
    }

    infra = softwareSystem "Infrastructure" {
      db = container "Postgres Database"
      cloud = container "Cloudinary"
    }

    user -> railsApp "Uploads video"
    railsApp -> fastapiService "POST /upload"
    fastapiService -> cloud "Stores video"
    fastapiService -> db "Writes guide data"
    railsApp -> db "Reads guide data"
  }

  views {
    systemContext rails {
      include *
      autoLayout
    }
    container rails {
      include *
      autoLayout
    }
  }
}
